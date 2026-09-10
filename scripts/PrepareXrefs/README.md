# XREF preparation

`removeXREFPaths.ps1` runs this worker automatically before transferring a selected
background to Xrefs. `-BindExplode 0` retains the previous transfer-only behavior.

The worker reads original DWGs without saving them, recursively prepares their
DWG dependencies, binds using separate symbol names, and explodes only the bound
references directly in modelspace. Ordinary blocks remain blocks; paperspace
references are bound but are not exploded. All intermediate files live in the
run's temporary folder. A completion marker gates transfer, so missing or failed
references cannot trigger archiving/replacement of the existing background.

ZIP extraction preserves directories and validates paths. Only selected DWGs
are delivered. Stored relative paths are tried first; stale paths can fall back
to an unambiguous filename in the source tree. ZIP dependencies must come from
the extracted archive, not another existing drawing on the computer.

Unloaded, missing, ambiguous, circular, clipped, or unbindable references fail
that drawing with a progress error. External images and underlays also require
manual preparation. These cases are not silently detached or exploded without
preserving their display. Other selected drawings continue processing.

Build from this repository with a .NET 8 SDK and installed AutoCAD 2022/2025 APIs:

```powershell
dotnet build scripts/PrepareXrefs/PrepareXrefs.csproj -c Release
```

The build stages both DLLs in `scripts/PrepareXrefs-bin`, included by the desktop
packaging spec. AutoCAD 2025 uses the .NET 8 DLL; earlier releases use .NET 4.8.

Live integration tests create disposable DWGs and require Core Console:

```powershell
$env:ACIES_TEST_CORE = 'C:/Program Files/Autodesk/AutoCAD 2025/accoreconsole.exe'
python -m unittest discover -s tests -p test_prepare_xrefs_core.py
```
