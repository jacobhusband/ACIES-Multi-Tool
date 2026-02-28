import shutil
import subprocess
import sys
import tempfile
import textwrap
import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
CAD_SCRIPT_PATHS = (
    REPO_ROOT / "scripts" / "PlotDWGs.ps1",
    REPO_ROOT / "scripts" / "FreezeLayersDWGs.ps1",
    REPO_ROOT / "scripts" / "ThawLayersDWGs.ps1",
)


class PowerShellCadWrapperTests(unittest.TestCase):
    def test_workroom_launch_context_uses_cached_cad_file_paths(self):
        text = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        self.assertIn("let workroomCadFilesLoading = false;", text)
        self.assertIn("let workroomCadFilesLoadPromise = Promise.resolve();", text)
        self.assertIn("let workroomDiscoveredCadFilePaths = [];", text)
        self.assertIn("function setWorkroomDiscoveredCadFilePaths(paths = []) {", text)
        self.assertIn("function setWorkroomCadFilesLoading(isLoading) {", text)
        self.assertIn("cadFilePaths: [...workroomDiscoveredCadFilePaths],", text)
        self.assertIn("setWorkroomDiscoveredCadFilePaths(discoveredPaths);", text)
        self.assertIn("setWorkroomDiscoveredCadFilePaths();", text)
        self.assertIn("async function triggerWorkroomTool(toolId) {", text)
        self.assertIn('message: "Waiting for CAD files..."', text)
        self.assertIn("await workroomCadFilesLoadPromise;", text)
        self.assertIn("disabled: isCadToolWaitingForLoad,", text)
        self.assertIn("cadFilesLoading: workroomCadFilesLoading,", text)
        self.assertIn("discoveredCadFileCount: workroomDiscoveredCadFilePaths.length,", text)

    def test_cad_scripts_use_raw_values_in_sta_relaunch(self):
        for script_path in CAD_SCRIPT_PATHS:
            with self.subTest(script=script_path.name):
                text = script_path.read_text(encoding="utf-8")
                self.assertNotIn('`"$PSCommandPath`"', text)
                self.assertNotIn('`"$FilesListPath`"', text)
                self.assertNotIn('`"$AcadCore`"', text)
                self.assertIn(
                    '@("-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-File", $PSCommandPath)',
                    text,
                )
                self.assertIn('@("-FilesListPath", $FilesListPath)', text)
                self.assertIn(
                    'Write-Host "PROGRESS: Received auto-selected files list: $FilesListPath"',
                    text,
                )

    @unittest.skipUnless(sys.platform == "win32", "PowerShell STA relaunch is Windows-only")
    def test_sta_relaunch_preserves_files_list_path(self):
        powershell = shutil.which("powershell.exe") or shutil.which("powershell")
        self.assertIsNotNone(powershell, "powershell.exe is required to validate the STA wrapper")

        harness = textwrap.dedent(
            """
            param(
              [string]$FilesListPath = ""
            )

            if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
              $ps = (Get-Process -Id $PID).Path
              $argsList = @("-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-File", $PSCommandPath)
              if ($PSBoundParameters.ContainsKey('FilesListPath') -and -not [string]::IsNullOrWhiteSpace($FilesListPath)) {
                $argsList += @("-FilesListPath", $FilesListPath)
              }
              $child = Start-Process -FilePath $ps -ArgumentList $argsList -Wait -PassThru
              exit $child.ExitCode
            }

            $files = @()
            $filesListWasProvided = $PSBoundParameters.ContainsKey('FilesListPath')
            $hasFilesListPath = $filesListWasProvided -and -not [string]::IsNullOrWhiteSpace($FilesListPath)
            if ($filesListWasProvided) {
              if ($hasFilesListPath) {
                Write-Host "PROGRESS: Received auto-selected files list: $FilesListPath"
                if (Test-Path $FilesListPath) {
                  $files = @(
                    Get-Content -Path $FilesListPath -Encoding UTF8 |
                      Where-Object { $_ -and $_.Trim() -and (Test-Path $_.Trim()) } |
                      ForEach-Object { $_.Trim() }
                  )
                  if ($files.Count -gt 0) {
                    Write-Host "PROGRESS: Using $($files.Count) DWG file(s) from auto-selected project folder."
                  }
                  else {
                    Write-Host "PROGRESS: Provided files list was empty. Opening file picker..."
                  }
                }
                else {
                  Write-Host "PROGRESS: Provided files list path was not found. Opening file picker..."
                }
              }
              else {
                Write-Host "PROGRESS: Files list parameter was provided without a path. Opening file picker..."
              }
            }

            if (-not $files -or $files.Count -eq 0) {
              Write-Host "PROGRESS: Waiting for user input..."
            }
            """
        ).strip()

        with tempfile.TemporaryDirectory(prefix="acies-ps-wrapper-") as temp_dir:
            temp_path = Path(temp_dir)
            dummy_dwg = temp_path / "fixture.dwg"
            dummy_dwg.write_text("", encoding="utf-8")
            files_list = temp_path / "files.txt"
            files_list.write_text(f"{dummy_dwg}\n", encoding="utf-8")
            harness_path = temp_path / "StaHarness.ps1"
            harness_path.write_text(harness, encoding="utf-8")

            result = subprocess.run(
                [
                    powershell,
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(harness_path),
                    "-FilesListPath",
                    str(files_list),
                ],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=30,
            )

            output = (result.stdout or "") + (result.stderr or "")
            self.assertEqual(0, result.returncode, msg=output)
            self.assertIn(
                f"PROGRESS: Received auto-selected files list: {files_list}",
                output,
            )
            self.assertIn(
                "PROGRESS: Using 1 DWG file(s) from auto-selected project folder.",
                output,
            )
            self.assertNotIn("PROGRESS: Waiting for user input...", output)


if __name__ == "__main__":
    unittest.main()
