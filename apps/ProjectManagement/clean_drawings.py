"""Headless cleanup orchestration. Originals are read/copied, never opened for saving."""
from datetime import datetime
import hashlib
import json
import math
import os
from pathlib import Path
import shutil
import stat
import subprocess
import tempfile
import threading
import time
import uuid

import fitz

_run_lock = threading.Lock()


def _file_hash(path):
    digest = hashlib.sha256()
    with Path(path).open('rb') as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b''):
            digest.update(chunk)
    return digest.hexdigest()


def _inside(path, root):
    try:
        Path(path).resolve().relative_to(Path(root).resolve())
        return True
    except ValueError:
        return False


def discover(project, notify=lambda message: None):
    notify('Finding electrical drawings and titleblock candidates…')
    root = Path(project).resolve(strict=True)
    electrical = next((p for p in root.iterdir() if p.is_dir() and p.name.lower() == 'electrical'), None)
    xrefs = [p for p in root.iterdir() if p.is_dir() and p.name.lower() in ('xref', 'xrefs')]
    drawings = sorted(p for p in electrical.glob('*') if p.is_file() and p.suffix.lower() == '.dwg') if electrical else []
    candidates = sorted(p for folder in xrefs for p in folder.rglob('*')
                        if p.is_file() and p.suffix.lower() == '.dwg' and _inside(p, root))
    candidates.sort(key=lambda p: (not any(h in p.stem.lower() for h in ('tblk', 'tblock', 'title', 'x-tb', 'border')), str(p)))
    pdfs = []
    notify('Finding published PDFs and checksets…')
    for folder in [root / 'PDF', (electrical / 'Checkset') if electrical else root / 'Electrical' / 'Checkset']:
        if folder.is_dir():
            pdfs.extend(p for p in folder.rglob('*') if p.is_file() and p.suffix.lower() == '.pdf' and _inside(p, root))
    sizes = []
    for pdf in sorted(pdfs, key=lambda p: p.stat().st_mtime, reverse=True)[:25]:
        notify(f'Reading PDF sheet sizes: {pdf.name}')
        try:
            if pdf.stat().st_size > 150 * 1024 * 1024:
                continue
            with fitz.open(pdf) as doc:
                seen = set()
                for index in range(min(len(doc), 100)):
                    rect = doc[index].rect  # Includes page rotation; preserve landscape/portrait.
                    width, height = round(rect.width / 72, 3), round(rect.height / 72, 3)
                    short, long = sorted((width, height))
                    if not any(abs(short - a) <= 0.5 and abs(long - b) <= 0.5
                               for a, b in ((22, 34), (24, 36), (30, 42), (36, 48))):
                        continue
                    if (width, height) in seen:
                        continue
                    seen.add((width, height))
                    sizes.append({'pdf': str(pdf.relative_to(root)), 'page': index + 1,
                                  'width': width, 'height': height})
        except Exception:
            continue
    return {'project': str(root), 'drawings': [str(p.relative_to(root)) for p in drawings],
            'titleblocks': [str(p.relative_to(root)) for p in candidates], 'sizes': sizes}


def plugin_for(acad):
    acad = Path(acad)
    if acad.name.lower() != 'accoreconsole.exe' or not acad.is_file():
        raise ValueError('Select an installed AutoCAD Core Console in settings.')
    # AutoCAD 2025+ uses .NET 8; earlier supported releases use .NET Framework.
    import re
    match = re.search(r'20\d{2}', str(acad.parent))
    if not match:
        raise ValueError('Cannot determine the AutoCAD release from its installation path.')
    framework = 'net8.0-windows' if int(match.group()) >= 2025 else 'net48'
    dll = Path(os.environ.get('APPDATA', '')) / 'Autodesk' / 'ApplicationPlugins' / 'ElectricalCommands.CleanCADCommands.bundle' / 'Contents' / framework / 'AutoCADCommands.CleanCADCommands.dll'
    if not dll.is_file():
        raise FileNotFoundError('Install the updated CleanCADCommands plugin before running Clean Drawings.')
    return dll


def preview(project, acad, notify=lambda message: None):
    """Read actual paper-space references before asking the user to choose."""
    result = discover(project, notify=notify)
    if not result['drawings'] or not result['titleblocks']:
        return result
    root = Path(result['project'])
    workspace = Path(tempfile.mkdtemp(prefix='acies-clean-preview-'))
    try:
        seed = workspace / 'seed.dwg'
        shutil.copy2(root / result['drawings'][0], seed)
        notify(f'Inspecting paper-space titleblock references in {len(result["drawings"])} drawing(s)…')
        scan = run_worker(acad, plugin_for(acad), seed,
                          {'Operation': 'scan', 'Files': [str(root / p) for p in result['drawings']],
                           'Catalog': [str(root / p) for p in result['titleblocks']]}, workspace, notify=notify)
    except Exception as exc:
        raise RuntimeError(f'{exc}\nInspection logs retained at: {workspace}') from exc
    else:
        try:
            shutil.rmtree(workspace)
        except OSError:
            result['inspectionWarning'] = f'Inspection files could not be removed: {workspace}'
    detected = {str(Path(p).resolve()).casefold() for p in scan.get('titleblocks', [])}
    result['detectedTitleblocks'] = [p for p in result['titleblocks'] if str((root / p).resolve()).casefold() in detected]
    result['titleblocks'].sort(key=lambda p: p not in result['detectedTitleblocks'])
    result['media'] = scan.get('media', [])
    return result


def run_worker(acad, dll, drawing, job, workspace, timeout=300, notify=lambda message: None):
    token = uuid.uuid4().hex
    folder = Path(workspace) / 'jobs' / token
    folder.mkdir(parents=True)
    request, result = folder / 'request.json', folder / 'result.json'
    job = dict(job, ResultPath=str(result))
    request.write_text(json.dumps(job), encoding='utf-8')
    script = folder / 'run.scr'
    # Paths are SCR arguments, not shell code. Reject line/quote injection.
    if any(c in str(dll) for c in '\r\n"'):
        raise ValueError('Invalid plugin path.')
    script.write_text('(setq aciesCleanSecure (getvar "SECURELOAD"))\n'
                      '(setvar "SECURELOAD" 0)\n_.NETLOAD\n"' + str(dll) + '"\n'
                      '(setvar "SECURELOAD" aciesCleanSecure)\n'
                      'ACIESCLEANJOB\n_.QUIT\n_Y\n', encoding='utf-8')
    env = dict(os.environ, ACIES_CLEAN_JOB=str(request))
    with (folder / 'worker.log').open('wb') as log:
        process = subprocess.Popen([str(acad), '/i', str(drawing), '/s', str(script), '/l', 'en-US'],
                                   cwd=str(folder), env=env, stdout=log, stderr=subprocess.STDOUT,
                                   creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
        started = time.monotonic()
        while True:
            remaining = timeout - (time.monotonic() - started)
            if remaining <= 0:
                process.kill()
                process.wait()
                raise RuntimeError(f'AutoCAD worker timed out; log: {folder / "worker.log"}')
            try:
                code = process.wait(timeout=min(5, remaining))
                break
            except subprocess.TimeoutExpired:
                elapsed = int(time.monotonic() - started)
                stage = {'scan': 'reference inspection', 'prepare': 'reference preparation',
                         'titleblock': 'titleblock cleanup', 'sheet': 'sheet cleanup',
                         'verify': 'saved drawing validation'}.get(job['Operation'], job['Operation'])
                notify(f'AutoCAD {stage}: running ({elapsed}s elapsed)…')
    if not result.is_file():
        raise RuntimeError(f'Worker returned no result (exit {code}). Update the plugin; log: {folder / "worker.log"}')
    response = json.loads(result.read_text(encoding='utf-8-sig'))
    if code != 0 or not response.get('success'):
        raise RuntimeError(response.get('error') or f'Worker exited with code {code}; log: {folder / "worker.log"}')
    return response.get('details', {})


def _selected(root, relative):
    path = (root / relative).resolve(strict=True)
    if not _inside(path, root) or path.suffix.lower() != '.dwg':
        raise ValueError('Select a DWG inside the project.')
    return path


def run(project, selection, acad, notify=lambda message: None, worker=run_worker, dll=None):
    if not _run_lock.acquire(blocking=False):
        raise RuntimeError('A Clean Drawings job is already running.')
    workspace = None
    if worker is run_worker:
        def worker(acad, dll, drawing, job, workspace):
            return run_worker(acad, dll, drawing, job, workspace, notify=notify)
    try:
        root = Path(project).resolve(strict=True)
        titleblock = _selected(root, selection['titleblock'])
        drawings = list(dict.fromkeys(_selected(root, p) for p in selection['drawings']))
        if not drawings or titleblock in drawings:
            raise ValueError('Select electrical drawings separately from the titleblock.')
        width, height = float(selection['width']), float(selection['height'])
        if not all(math.isfinite(v) and 0 < v <= 1000 for v in (width, height)):
            raise ValueError('Enter valid sheet dimensions in CAD units.')
        dll = dll or plugin_for(acad)
        workspace = Path(tempfile.mkdtemp(prefix='acies-clean-'))
        seed = workspace / 'seed.dwg'
        shutil.copy2(titleblock, seed)
        seed.chmod(stat.S_IREAD | stat.S_IWRITE)
        notify('Inspecting drawings and required XREFs…')
        catalog = [str(root / p) for p in discover(root)['titleblocks']]
        scan = worker(acad, dll, seed, {'Operation': 'scan', 'Files': [str(titleblock), *map(str, drawings)], 'Catalog': catalog}, workspace)
        if scan.get('media'):
            raise RuntimeError('Headless image/underlay/linked OLE embedding is not supported yet. No drawings were cleaned.\n' + '\n'.join(scan['media'][:12]))
        copies = {}
        fingerprints = {}
        notify(f'Copying {len(scan["files"])} drawing dependencies to local working storage…')
        for source in scan['files']:
            path = Path(source).resolve(strict=True)
            relative = path.relative_to(root) if _inside(path, root) else Path('Dependencies') / hashlib.sha256(str(path).encode()).hexdigest()[:12] / path.name
            target = workspace / 'project' / relative
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(path, target)
            target.chmod(stat.S_IREAD | stat.S_IWRITE)
            fingerprints[str(path)] = _file_hash(target)
            copies[str(path)] = str(target)
        notify('Redirecting references to the local working copies…')
        worker(acad, dll, seed, {'Operation': 'prepare', 'Copies': copies, 'References': scan['references']}, workspace)
        staged_titleblock = Path(copies[str(titleblock)])
        output_root = workspace / 'outputs'
        output_root.mkdir()
        cleaned_titleblock = output_root / (uuid.uuid4().hex + '.dwg')
        notify('Cleaning the titleblock…')
        title_result = worker(acad, dll, staged_titleblock, {'Operation': 'titleblock', 'Width': width, 'Height': height, 'Output': str(cleaned_titleblock)}, workspace)
        notify('Validating the saved titleblock…')
        reopened = worker(acad, dll, cleaned_titleblock, {'Operation': 'verify'}, workspace)
        if reopened != title_result:
            raise RuntimeError('Titleblock entity counts changed after saving and reopening.')
        shutil.copy2(cleaned_titleblock, staged_titleblock)
        results = []
        outputs = [(staged_titleblock, titleblock.relative_to(root))]
        for index, drawing in enumerate(drawings, 1):
            notify(f'Cleaning drawing {index} of {len(drawings)}: {drawing.name}')
            staged = Path(copies[str(drawing)])
            output = output_root / (uuid.uuid4().hex + '.dwg')
            counts = worker(acad, dll, staged, {'Operation': 'sheet', 'Titleblock': str(staged_titleblock),
                                              'Width': width, 'Height': height, 'Output': str(output)}, workspace)
            notify(f'Validating drawing {index} of {len(drawings)}: {drawing.name}')
            reopened = worker(acad, dll, output, {'Operation': 'verify'}, workspace)
            if reopened != counts:
                raise RuntimeError(f'Entity counts changed after saving and reopening {drawing.name}.')
            outputs.append((output, drawing.relative_to(root)))
            results.append({'drawing': str(drawing.relative_to(root)), 'validation': reopened})
        notify('Checking source versions and delivering validated drawings…')
        for source, digest in fingerprints.items():
            if _file_hash(source) != digest:
                raise RuntimeError(f'Source changed during processing. Run again: {source}')
        report = {'status': 'success', 'selection': selection, 'drawings': results,
                  'titleblockValidation': title_result, 'sourceHashes': fingerprints,
                  'limitations': ['External images, underlays, and linked OLE are unsupported.',
                                  'Boundary intersection uses conservative entity bounds.',
                                  'Reopen/count/reference checks are not a visual comparison.']}
        destination_root = root / 'Cleaned CAD'
        if not _inside(destination_root, root):
            raise RuntimeError('Cleaned CAD resolves outside the project directory.')
        destination_root.mkdir(exist_ok=True)
        stamp = datetime.now().strftime('%Y-%m-%d_%H%M%S') + '_' + uuid.uuid4().hex[:6]
        pending = destination_root / ('.incomplete-' + stamp)
        pending.mkdir()
        for source, relative in outputs:
            destination = pending / relative
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source, destination)
            if _file_hash(source) != _file_hash(destination):
                raise RuntimeError(f'Output copy verification failed: {destination}')
        (pending / 'cleanup-report.json').write_text(json.dumps(report, indent=2), encoding='utf-8')
        destination = destination_root / stamp
        pending.rename(destination)
        # Remove only the unique local directory created by this invocation.
        cleanup_warning = ''
        try:
            shutil.rmtree(workspace)
            workspace = None
        except OSError:
            cleanup_warning = f'Drawings delivered successfully. Temporary files could not be removed: {workspace}'
        return {'status': 'success', 'output': str(destination), 'count': len(drawings), 'cleanupWarning': cleanup_warning}
    except Exception as exc:
        raise RuntimeError(str(exc) + (f'\nWorking files and logs retained at: {workspace}' if workspace else '')) from exc
    finally:
        _run_lock.release()
