import shutil
import subprocess
import sys
import tempfile
import textwrap
import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
MAIN_PY_PATH = REPO_ROOT / "main.py"
SCRIPT_JS_PATH = REPO_ROOT / "script.js"
CAD_SCRIPT_PATHS = (
    REPO_ROOT / "scripts" / "PlotDWGs.ps1",
    REPO_ROOT / "scripts" / "FreezeLayersDWGs.ps1",
    REPO_ROOT / "scripts" / "ThawLayersDWGs.ps1",
)


class PowerShellCadWrapperTests(unittest.TestCase):
    def test_backend_trace_messages_are_logged_but_not_forwarded_to_ui(self):
        text = MAIN_PY_PATH.read_text(encoding="utf-8")
        self.assertIn('if message.startswith("TRACE"):', text)
        self.assertIn("'script_trace'", text)
        self.assertIn("continue", text)

    def test_frontend_automation_keeps_cad_trace_helpers_without_workroom_modal(self):
        text = SCRIPT_JS_PATH.read_text(encoding="utf-8")
        self.assertIn("async getCadAutoSelectTrace(lineLimit = 200) {", text)
        self.assertIn("async clearCadAutoSelectTrace() {", text)
        self.assertIn("window.pywebview?.api?.get_cad_auto_select_trace", text)
        self.assertIn("window.pywebview?.api?.clear_cad_auto_select_trace", text)
        self.assertNotIn("openWorkroom(projectIndex = 0)", text)
        self.assertNotIn("async runWorkroomTool(toolId) {", text)
        self.assertNotIn("getWorkroomState()", text)
        self.assertNotIn("getToolState(toolId)", text)

    def test_frontend_restores_workroom_launch_context_helpers_for_shared_tools(self):
        text = SCRIPT_JS_PATH.read_text(encoding="utf-8")

        for expected in (
            "function getActiveWorkroomContext() {",
            "function consumePendingCadLaunchContext() {",
            "function buildWorkroomCadLaunchContext() {",
            "function resolveCadLaunchContextForTool() {",
            'source: "workroom",',
            "rootProjectPath,",
            "cadFilePaths: [],",
            "projectName:",
            "deliverableName:",
            "const pendingContext = consumePendingCadLaunchContext();",
            'const checklistModal = document.getElementById("checklistModal");',
            "if (!checklistModal?.open) return null;",
        ):
            self.assertIn(expected, text)

    def test_cad_scripts_use_raw_values_in_sta_relaunch(self):
        for script_path in CAD_SCRIPT_PATHS:
            with self.subTest(script=script_path.name):
                text = script_path.read_text(encoding="utf-8")
                self.assertIn("function Move-FormToPrimaryScreen {", text)
                self.assertIn("[System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea", text)
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
                self.assertIn(
                    'Write-Host "PROGRESS: TRACE files_list_param_bound=$([int]$filesListWasProvided) path=$FilesListPath"',
                    text,
                )
                self.assertIn(
                    'Write-Host "PROGRESS: TRACE branch=auto_selected_files count=$($files.Count)"',
                    text,
                )
                self.assertIn(
                    'Write-Host "PROGRESS: TRACE branch=manual_picker"',
                    text,
                )
                if script_path.name == "PlotDWGs.ps1":
                    self.assertIn('$form.StartPosition = "Manual"', text)
                    self.assertIn('$form.TopMost = $true', text)
                    self.assertIn('$form.ShowInTaskbar = $true', text)
                    self.assertIn(
                        '$form.WindowState = [System.Windows.Forms.FormWindowState]::Normal',
                        text,
                    )
                    self.assertIn("Move-FormToPrimaryScreen $form", text)
                    self.assertIn('$form.BringToFront()', text)
                    self.assertIn(
                        'Write-Host "PROGRESS: Waiting for paper size confirmation..."',
                        text,
                    )
                    self.assertIn(
                        'Write-Host "PROGRESS: Paper size dialog should be visible on the primary display."',
                        text,
                    )
                    self.assertIn(
                        'Write-Host "PROGRESS: TRACE branch=paper_size_dialog"',
                        text,
                    )
                else:
                    self.assertIn('$form.StartPosition = "Manual"', text)
                    self.assertIn("Move-FormToPrimaryScreen $form", text)
                    self.assertIn(
                        '$form.WindowState = [System.Windows.Forms.FormWindowState]::Normal',
                        text,
                    )
                    self.assertIn('Write-Host "PROGRESS: Reading extracted data..."', text)
                    self.assertIn('$form.TopMost = $true', text)
                    self.assertIn('$form.ShowInTaskbar = $true', text)
                    self.assertIn('$form.BringToFront()', text)
                    self.assertIn(
                        'Write-Host "PROGRESS: Waiting for layer selection..."',
                        text,
                    )
                    self.assertIn(
                        'Write-Host "PROGRESS: Layer selection dialog should be visible on the primary display."',
                        text,
                    )
                    self.assertIn(
                        'Write-Host "PROGRESS: TRACE branch=layer_selection_dialog"',
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
