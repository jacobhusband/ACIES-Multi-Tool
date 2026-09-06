# Clean Drawings

Launch **Clean Drawings** from the Workroom or General tools with a project selected.
Install the updated `ElectricalCommands.CleanCADCommands` bundle first and select
AutoCAD Core Console in settings. The existing desktop commands remain available.

The confirmation screen lists DWGs directly in `Electrical`, titleblock candidates
under `XREF` or `Xrefs`, and standard sheet sizes found in recent PDFs under `PDF`
and `Electrical/Checkset`. Candidates actually referenced in paper space are marked
and ranked first. Confirm the correct XREF, drawing files, and dimensions. PDF
dimensions are inches; the editable dimensions are CAD units. Mixed PDF page sizes
are offered separately, preserving page orientation. No usable PDF requires manual
dimensions; cleanup never guesses a size from drawing extents.

The titleblock source boundary is `(0,0)` to `(width,height)` in Model Space.
The same rectangle is transformed by each titleblock insertion for paper-space
cleanup. Entity bounding-box intersection is conservative: crossing entities are
retained whole, and a bounding box may intersect even if the entity itself does not.

## Processing and output

1. Inspect the selected DWGs and recursively resolve their DWG dependencies.
   A missing path can resolve to a unique filename under the project's XREF tree;
   missing or ambiguous dependencies stop processing.
2. Copy the dependency graph to a unique local temporary directory. Preserve
   project-relative structure and place outside-project dependencies in unique
   `Dependencies` directories. Rewrite references in these copies only.
3. Clean the titleblock, save to a separate temporary output, and reopen it.
4. Replace the staged titleblock with its validated result. Clean each sheet using
   the explicitly confirmed XREF identity and synchronous `CLEANCAD2` stages.
5. Check for empty drawings, remaining external DWG references, external media,
   and changes in entity counts after reopening. Check source hashes before delivery.
6. Deliver all results together to
   `Project/Cleaned CAD/YYYY-MM-DD_HHMMSS_<unique>/`, preserving `Electrical` and
   XREF paths, with `cleanup-report.json`. The report records confirmed dimensions,
   PDF provenance, validation results, and source hashes.

Original drawings are never saved. A failure retains local working files and logs,
reports their location, and does not deliver a partial successful batch. A failure
during the final copy can leave a `.incomplete-*` directory in `Cleaned CAD`.
The tool allows one cleanup batch at a time per app process. Each worker has a
five-minute timeout. Cancellation during processing is not implemented in this version.

## Current limits

- Wipeout masks are supported; they are not treated as external image attachments.
  Validation records their count and checks it again after reopening the output.
- External raster images, PDF/DGN/DWF underlays, and linked OLE objects stop the batch before cleanup.
  The desktop PowerPoint/clipboard/OLE conversion has no headless replacement yet.
- The titleblock must contain measurable local geometry. Local blocks that cannot
  be exploded stop processing. Its DWG XREFs are intentionally detached.
- Each sheet must have a uniquely identified titleblock definition in paper space
  and usable viewports for Model Space cleanup. Multiple titleblock insertions on
  one layout currently stop automatic paper-space cleanup.
- Reopen, count, and reference checks are structural validation, not a rendered
  visual comparison or a guarantee about embedded OLE display/printing.
- There is no full-AutoCAD fallback: unsupported drawings are reported explicitly.

## Batch contract

`ACIES_CLEAN_JOB` is a process-local environment variable containing the absolute
path of a UTF-8 JSON request. Run `ACIESCLEANJOB` after `NETLOAD`. Requests contain
`Operation` and `ResultPath`, plus operation-specific fields:

| Operation | Fields | Purpose |
| --- | --- | --- |
| `scan` | `Files`, `Catalog` | Read dependency graph, paper-space XREF candidates, media |
| `prepare` | `Copies`, `References` | Rewrite staged DWG reference paths |
| `titleblock` | `Width`, `Height`, `Output` | Fixed-origin titleblock cleanup |
| `sheet` | `Titleblock`, `Width`, `Height`, `Output` | Confirmed-titleblock sheet cleanup |
| `verify` | none | Inspect active saved output |
| `embed-prototype` | `Output` (new file only) | Experimental raster-to-native-solids conversion, separate from cleanup |

`CLEANTBLK2` is an entry point for a `titleblock` request using this same contract.
Each worker writes `{ "success": true, "details": ... }` or
`{ "success": false, "error": ... }`. The app requires both a successful process
exit and a successful result; a log marker alone cannot authorize delivery.

Run unit tests with `.venv/Scripts/python.exe -m unittest tests.test_clean_drawings`.
The optional real AutoCAD test `tests.test_clean_drawings_core` requires
`ACIES_TEST_CORE` and `ACIES_TEST_CLEAN_DLL` pointing to the matching installed
Core Console and built plugin DLL. It creates disposable synthetic DWGs and
checks fixed-origin pruning, XREF binding, reopen results, and original preservation.

## Experimental self-contained image geometry

`embed-prototype` replaces ordinary attached raster images with blocks containing
true-color `SOLID` entities. Recipients need neither the original image files nor
the ACIES plugin. This is an explicit worker operation; the Clean Drawings button
still rejects external media. PDF/DGN/DWF underlays and linked OLE are not converted.

The prototype decodes source pixels without downsampling or color quantization,
merges equal horizontal runs and matching runs on adjacent rows, and maps them
through the image's original origin and width/height vectors. It preserves the
owner block, layer, entity properties and position in draw order. Unreferenced
raster image definitions are removed; wipeouts are excluded from conversion.

Supported inputs are opaque, non-bitonal images parallel to their owner XY plane,
with default brightness/contrast, no fade, image transparency off, unlocked layers
and no partial clipping. AutoCAD's default whole-image clip is accepted. FILLMODE
must be 1. The source path must exist exactly as stored (relative to the input DWG
when applicable); source dimensions must match the attachment. Limits are two
million pixels per image and 50,000 generated solids per drawing. All images are
validated before mutation. A rejection produces no output DWG. The output path
must not already exist. Use disposable local fixtures while evaluating this API.

Invoke using the existing worker from the ProjectManagement Python environment:

```python
from pathlib import Path
from clean_drawings import run_worker

result = run_worker(core_console, plugin_dll, input_dwg,
                    {'Operation': 'embed-prototype', 'Output': str(output_dwg)},
                    Path(local_job_directory))
assert run_worker(core_console, plugin_dll, output_dwg,
                  {'Operation': 'verify'}, Path(local_job_directory)) == result['validation']
```

The result includes image/solid counts under `embedding` and structural checks
under `validation`. Solids follow CAD fill, viewport, and plot-style behavior;
they are not embedded bitmap/OLE objects. Photographs and anti-aliased logos can
generate many entities. Nested transforms, arbitrary plot styles and real project
logos need broader visual coverage before enabling this in the delivery pipeline.

Run `tests.test_native_image_geometry_core` with `ACIES_TEST_CORE` and
`ACIES_TEST_CLEAN_DLL`. It tests rotated pixel geometry after deleting the PNG,
reopens with built-in AutoLISP without NETLOAD, checks native colors/coordinates,
and rejects alpha, partial clipping and excessive geometry. Set
`ACIES_NATIVE_PLOT=1` to plot the source-free output to PDF and verify its vector
fill colors. `ACIES_NATIVE_ARTIFACTS` optionally retains artifacts in a new directory.
