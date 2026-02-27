import argparse
import json
import os
import shutil
import sys
import tempfile
import time
import traceback
from datetime import datetime
from pathlib import Path

import webview
from PIL import ImageGrab


REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

# Ensure deterministic CAD tool behavior in the backend for automation runs.
os.environ.setdefault("ACIES_TEST_MODE", "1")

from main import Api, BASE_DIR, SETTINGS_FILE, TASKS_FILE  # noqa: E402


TEST_MODE_TOOL_IDS = [
    "toolPublishDwgs",
    "toolFreezeLayers",
    "toolThawLayers",
    "toolCleanXrefs",
]
WORKROOM_TOOL_IDS = TEST_MODE_TOOL_IDS + [
    "toolCreateNarrativeTemplate",
    "toolCreatePlanCheckTemplate",
    "toolLightingSchedule",
    "toolTitle24Compliance",
]
DIRECT_CARD_LAUNCH_TOOL_IDS = {"toolPublishDwgs"}
WORKROOM_PROJECT_MODAL_TOOL_IDS = {
    "toolLightingSchedule": {
        "dialog_id": "lightingScheduleDlg",
        "select_id": "lightingScheduleProjectSelect",
    },
    "toolTitle24Compliance": {
        "dialog_id": "title24ComplianceDlg",
        "select_id": "title24ProjectSelect",
    },
}


class BackupFile:
    def __init__(self, file_path):
        self.file_path = Path(file_path)
        self._temp_copy = None
        self._did_exist = False

    def __enter__(self):
        self._did_exist = self.file_path.exists()
        if self._did_exist:
            fd, tmp_path = tempfile.mkstemp(suffix=".bak")
            os.close(fd)
            self._temp_copy = Path(tmp_path)
            shutil.copy2(self.file_path, self._temp_copy)
        return self

    def __exit__(self, exc_type, exc, tb):
        if self._did_exist and self._temp_copy and self._temp_copy.exists():
            self.file_path.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(self._temp_copy, self.file_path)
            self._temp_copy.unlink(missing_ok=True)
        elif not self._did_exist and self.file_path.exists():
            self.file_path.unlink()
        return False


def write_json(path, payload):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)


def read_json_or_default(path, default):
    path = Path(path)
    if not path.exists():
        return default
    try:
        with path.open("r", encoding="utf-8") as handle:
            data = json.load(handle)
            return data if isinstance(data, type(default)) else default
    except Exception:
        return default


def build_fixture_settings(existing=None):
    settings = dict(existing or {})
    settings["discipline"] = ["Electrical"]
    settings["workroomAutoSelectCadFiles"] = True
    if not settings.get("autocadPath"):
        settings["autocadPath"] = r"C:\Program Files\Autodesk\AutoCAD 2025\accoreconsole.exe"
    settings.setdefault("publishDwgOptions", {"autoDetectPaperSize": True, "shrinkPercent": 100})
    settings.setdefault("freezeLayerOptions", {"scanAllLayers": True})
    settings.setdefault("thawLayerOptions", {"scanAllLayers": True})
    settings.setdefault("cleanDwgOptions", {
        "stripXrefs": True,
        "setByLayer": True,
        "purge": True,
        "audit": True,
        "hatchColor": True,
    })
    return settings


def build_fixture_tasks(saved_project_path):
    deliverable_id = "dlv_e2e_workroom"
    return [
        {
            "id": "E2E-1001",
            "name": "E2E Workroom Validation Project",
            "nick": "",
            "notes": "",
            "path": str(saved_project_path),
            "refs": [],
            "deliverables": [
                {
                    "id": deliverable_id,
                    "name": "E2E Deliverable",
                    "due": "",
                    "notes": "",
                    "tasks": [{"text": "Validate Workroom CAD auto-select", "done": False, "links": []}],
                    "statuses": ["Working"],
                    "statusTags": ["working"],
                    "status": "Working",
                    "emailRefs": [],
                    "emailRef": None,
                    "workroomCadDiscipline": "Electrical",
                }
            ],
            "overviewDeliverableId": deliverable_id,
        }
    ]


def parse_js_json(window, expression):
    raw = window.evaluate_js(f"JSON.stringify(({expression}))")
    if raw in (None, "", "undefined", "null"):
        return None
    if isinstance(raw, str):
        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            return raw
    return raw


def wait_for(predicate, timeout_seconds, interval_seconds=0.2):
    start = time.time()
    last_value = None
    while time.time() - start <= timeout_seconds:
        last_value = predicate()
        if last_value:
            return last_value
        time.sleep(interval_seconds)
    raise TimeoutError(f"Timed out after {timeout_seconds}s")


def normalize_path_for_compare(path):
    return os.path.normcase(os.path.normpath(str(path or "")))


def is_subpath_or_equal(path, root):
    normalized_path = normalize_path_for_compare(path).rstrip("\\/")
    normalized_root = normalize_path_for_compare(root).rstrip("\\/")
    if not normalized_path or not normalized_root:
        return False
    return (
        normalized_path == normalized_root
        or normalized_path.startswith(normalized_root + os.sep)
    )


def capture_window_screenshot(window, output_path):
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        x = int(getattr(window, "x"))
        y = int(getattr(window, "y"))
        width = int(getattr(window, "width"))
        height = int(getattr(window, "height"))
        bbox = (max(0, x), max(0, y), max(1, x + width), max(1, y + height))
        image = ImageGrab.grab(bbox=bbox)
    except Exception:
        image = ImageGrab.grab()
    image.save(output_path)
    return str(output_path)


def wait_for_tool_phase(window, tool_id, timeout_seconds):
    def _poll():
        state = parse_js_json(
            window,
            f"window.__aciesAutomation && window.__aciesAutomation.getToolState('{tool_id}')",
        )
        if not isinstance(state, dict):
            return None
        if state.get("phase") in ("done", "error"):
            return state
        return None

    return wait_for(_poll, timeout_seconds)


def collect_workroom_icon_state(window, tool_ids):
    serialized_tool_ids = json.dumps(tool_ids)
    return parse_js_json(
        window,
        f"""(() => {{
            const ids = {serialized_tool_ids};
            return ids.map((toolId) => {{
                const row = document.querySelector(`#workroomToolsList .workroom-tool-item[data-tool-id="${{toolId}}"]`);
                const icon = row ? row.querySelector(".workroom-tool-icon") : null;
                return {{
                    toolId,
                    rowFound: !!row,
                    hasIcon: !!icon,
                    hasSvg: !!icon?.querySelector("svg"),
                    iconText: (icon?.textContent || "").trim(),
                }};
            }});
        }})()""",
    ) or []


def collect_workroom_discipline_state(window):
    return parse_js_json(
        window,
        """(() => {
            const container = document.getElementById("workroomCadRouting");
            if (!container) {
                return { exists: false, hidden: true, options: [] };
            }
            const options = Array.from(
                container.querySelectorAll('input[name="workroomCadDiscipline"]')
            ).map((input) => ({
                value: String(input.value || ""),
                checked: !!input.checked,
            }));
            return {
                exists: true,
                hidden: !!container.hidden,
                options,
            };
        })()""",
    ) or {"exists": False, "hidden": True, "options": []}


def run_tool_from_direct_card_click(window, tool_id):
    return parse_js_json(
        window,
        f"""(() => {{
            const card = document.getElementById("{tool_id}");
            if (!card) {{
                return {{ toolId: "{tool_id}", exists: false }};
            }}
            card.click();
            return {{
                toolId: "{tool_id}",
                exists: true,
                state: window.__aciesAutomation?.getToolState("{tool_id}") || null,
            }};
        }})()""",
    ) or {}


def wait_for_workroom_project_modal(window, tool_id, timeout_seconds):
    modal_config = WORKROOM_PROJECT_MODAL_TOOL_IDS.get(tool_id)
    if not modal_config:
        return None
    dialog_id = modal_config["dialog_id"]
    select_id = modal_config["select_id"]

    def _poll():
        state = parse_js_json(
            window,
            f"""(() => {{
                const dlg = document.getElementById("{dialog_id}");
                const select = document.getElementById("{select_id}");
                return {{
                    open: !!dlg?.open,
                    selected: select?.value ?? "",
                }};
            }})()""",
        )
        if isinstance(state, dict) and state.get("open"):
            return state
        return None

    return wait_for(_poll, timeout_seconds)


def close_modal_by_id(window, dialog_id):
    parse_js_json(
        window,
        f"""(() => {{
            const dlg = document.getElementById("{dialog_id}");
            if (dlg?.open) dlg.close();
            return {{ open: !!dlg?.open }};
        }})()""",
    )


def run_automation(
    window,
    api,
    artifacts_dir,
    project_root,
    saved_project_path,
    timeout_seconds,
    result_holder,
):
    screenshots = []
    tool_results = []
    try:
        expected_saved_project_path = normalize_path_for_compare(saved_project_path)
        expected_mechanical_folder = normalize_path_for_compare(
            project_root / "Mechanical"
        )
        expected_arch_folder = normalize_path_for_compare(project_root / "Arch")

        wait_for(
            lambda: parse_js_json(
                window,
                "document.readyState === 'complete' && !!window.__aciesAutomation",
            ),
            timeout_seconds=timeout_seconds,
            interval_seconds=0.25,
        )

        workroom_state = parse_js_json(
            window, "window.__aciesAutomation.openWorkroom(0)"
        ) or {}
        if not workroom_state.get("modalOpen"):
            raise RuntimeError("Failed to open Project Workroom modal for automation.")

        screenshots.append(
            capture_window_screenshot(window, artifacts_dir / "01-workroom-open.png")
        )
        workroom_icon_state = collect_workroom_icon_state(window, WORKROOM_TOOL_IDS)
        discipline_state = collect_workroom_discipline_state(window)
        if not discipline_state.get("exists") or discipline_state.get("hidden"):
            raise RuntimeError("Workroom discipline selector is missing or hidden.")
        discipline_options = discipline_state.get("options") or []
        if not discipline_options:
            raise RuntimeError("Workroom discipline selector rendered without options.")
        checked_values = [
            entry.get("value")
            for entry in discipline_options
            if entry.get("checked")
        ]
        if checked_values != ["Electrical"]:
            raise RuntimeError(
                "Workroom discipline selector did not default to Electrical in single-discipline mode "
                f"(checked={checked_values})."
            )
        selected_discipline_state = parse_js_json(
            window,
            """(() => {
                const target = document.querySelector(
                    'input[name="workroomCadDiscipline"][value="Mechanical"]'
                );
                if (!target) {
                    return { exists: false, selected: "" };
                }
                target.click();
                return {
                    exists: true,
                    selected: target.checked ? "Mechanical" : "",
                };
            })()""",
        ) or {}
        if not selected_discipline_state.get("exists"):
            raise RuntimeError(
                "Workroom discipline selector is missing the Mechanical option."
            )
        discipline_state = collect_workroom_discipline_state(window)
        selected_checked_values = [
            entry.get("value")
            for entry in (discipline_state.get("options") or [])
            if entry.get("checked")
        ]
        if selected_checked_values != ["Mechanical"]:
            raise RuntimeError(
                "Workroom discipline selector did not persist Mechanical selection "
                f"(checked={selected_checked_values})."
            )
        missing_workroom_rows = [
            entry.get("toolId")
            for entry in workroom_icon_state
            if not entry.get("rowFound") or not entry.get("hasIcon")
        ]
        if missing_workroom_rows:
            raise RuntimeError(
                "Missing workroom tool rows/icons for: "
                + ", ".join(missing_workroom_rows)
            )
        missing_svg_icons = [
            entry.get("toolId")
            for entry in workroom_icon_state
            if not entry.get("hasSvg")
        ]
        if missing_svg_icons:
            raise RuntimeError(
                "Workroom tools rendered without SVG icons: "
                + ", ".join(missing_svg_icons)
            )

        for index, tool_id in enumerate(WORKROOM_TOOL_IDS, start=2):
            if tool_id in DIRECT_CARD_LAUNCH_TOOL_IDS:
                direct_run_state = run_tool_from_direct_card_click(window, tool_id)
                if not direct_run_state.get("exists"):
                    raise RuntimeError(f"Missing source tool card for direct launch: {tool_id}")
            else:
                parse_js_json(window, f"window.__aciesAutomation.runWorkroomTool('{tool_id}')")
            if tool_id in WORKROOM_PROJECT_MODAL_TOOL_IDS:
                modal_state = wait_for_workroom_project_modal(
                    window, tool_id, timeout_seconds=timeout_seconds
                )
                selected_project = str(modal_state.get("selected") or "").strip()
                if selected_project != "0":
                    raise RuntimeError(
                        f"{tool_id} did not auto-select the active Workroom project "
                        f"(selected={selected_project or '<empty>'})."
                    )
                dialog_id = WORKROOM_PROJECT_MODAL_TOOL_IDS[tool_id]["dialog_id"]
                close_modal_by_id(window, dialog_id)
                state = {
                    "toolId": tool_id,
                    "exists": True,
                    "running": False,
                    "phase": "done",
                    "statusText": "DONE",
                }
            else:
                state = wait_for_tool_phase(window, tool_id, timeout_seconds=timeout_seconds)
            shot_path = artifacts_dir / f"{index:02d}-{tool_id}.png"
            screenshots.append(capture_window_screenshot(window, shot_path))
            launch_mode = "direct_card_click" if tool_id in DIRECT_CARD_LAUNCH_TOOL_IDS else "workroom_run_button"
            tool_results.append({**state, "launchMode": launch_mode})
            if state.get("phase") != "done":
                raise RuntimeError(
                    f"{tool_id} failed during automation. status={state.get('statusText')}"
                )

        narrative_outputs = sorted(saved_project_path.glob("Narrative of Changes*.docx"))
        if not narrative_outputs:
            raise RuntimeError(
                "Narrative template was not auto-created in the saved project path."
            )
        plan_check_outputs = sorted(saved_project_path.glob("Plan Check Comments*.doc"))
        if not plan_check_outputs:
            raise RuntimeError(
                "Plan Check template was not auto-created in the saved project path."
            )

        test_records = api.get_test_mode_records()
        record_list = test_records.get("records", [])
        successful_tools = {
            record.get("tool_id")
            for record in record_list
            if record.get("status") == "success"
        }
        missing = [
            tool_id for tool_id in TEST_MODE_TOOL_IDS if tool_id not in successful_tools
        ]
        if missing:
            raise RuntimeError(
                f"Missing successful backend test records for: {', '.join(missing)}"
            )

        for cad_tool_id in ("toolPublishDwgs", "toolFreezeLayers", "toolThawLayers"):
            record = next(
                (
                    item
                    for item in reversed(record_list)
                    if item.get("tool_id") == cad_tool_id and item.get("status") == "success"
                ),
                None,
            )
            if not record:
                raise RuntimeError(f"Missing successful backend test record for {cad_tool_id}.")
            source = str(record.get("source") or "").strip().lower()
            if source != "workroom":
                raise RuntimeError(
                    f"{cad_tool_id} test record did not keep workroom launch source (source={source or '<empty>'})."
                )
            project_path = normalize_path_for_compare(record.get("project_path"))
            if project_path != expected_saved_project_path:
                raise RuntimeError(
                    f"{cad_tool_id} did not keep the nested saved project path in launch context "
                    f"(project_path={record.get('project_path') or '<empty>'})."
                )
            if int(record.get("count") or 0) <= 0:
                raise RuntimeError(
                    f"{cad_tool_id} did not report auto-selected DWGs in workroom test mode."
                )
            resolved_discipline = str(record.get("discipline") or "").strip()
            if resolved_discipline != "Mechanical":
                raise RuntimeError(
                    f"{cad_tool_id} did not keep the selected Mechanical discipline "
                    f"(discipline={resolved_discipline or '<empty>'})."
                )
            folder_path = normalize_path_for_compare(record.get("folder_path"))
            if folder_path != expected_mechanical_folder:
                raise RuntimeError(
                    f"{cad_tool_id} resolved DWGs from an unexpected folder "
                    f"(folder_path={record.get('folder_path') or '<empty>'})."
                )
            if is_subpath_or_equal(folder_path, saved_project_path):
                raise RuntimeError(
                    f"{cad_tool_id} incorrectly resolved DWGs inside the nested saved path "
                    f"(folder_path={record.get('folder_path') or '<empty>'})."
                )

        clean_record = next(
            (
                record
                for record in record_list
                if record.get("tool_id") == "toolCleanXrefs"
                and record.get("status") == "success"
            ),
            None,
        )
        if not clean_record:
            raise RuntimeError("Missing successful backend test record for toolCleanXrefs.")
        if not clean_record.get("manual_selection_enforced"):
            raise RuntimeError(
                "toolCleanXrefs did not enforce manual file selection in workroom test mode."
            )
        if not clean_record.get("arch_folder"):
            raise RuntimeError(
                "toolCleanXrefs did not resolve an Arch folder for workroom manual selection."
            )
        clean_project_path = normalize_path_for_compare(clean_record.get("project_path"))
        if clean_project_path != expected_saved_project_path:
            raise RuntimeError(
                "toolCleanXrefs did not keep the nested saved project path in launch context."
            )
        clean_arch_folder = normalize_path_for_compare(clean_record.get("arch_folder"))
        if clean_arch_folder != expected_arch_folder:
            raise RuntimeError(
                "toolCleanXrefs did not resolve the Arch folder from the 6-digit project root."
            )

        result_holder["status"] = "success"
        result_holder["workroom_state"] = workroom_state
        result_holder["workroom_icons"] = workroom_icon_state
        result_holder["discipline_state"] = discipline_state
        result_holder["tools"] = tool_results
        result_holder["screenshots"] = screenshots
        result_holder["test_records"] = test_records
        result_holder["template_outputs"] = {
            "narrative": [str(path) for path in narrative_outputs],
            "planCheck": [str(path) for path in plan_check_outputs],
        }
    except Exception as exc:
        result_holder["status"] = "error"
        result_holder["message"] = str(exc)
        result_holder["traceback"] = traceback.format_exc()
    finally:
        try:
            window.destroy()
        except Exception:
            pass


def main():
    parser = argparse.ArgumentParser(
        description="Run automated E2E validation for Project Workroom CAD tools."
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=45,
        help="Per-step timeout in seconds.",
    )
    parser.add_argument(
        "--keep-fixtures",
        action="store_true",
        help="Keep generated fixture files after run.",
    )
    args = parser.parse_args()

    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    artifacts_dir = REPO_ROOT / "build" / "e2e" / "workroom" / timestamp
    artifacts_dir.mkdir(parents=True, exist_ok=True)

    fixture_root = Path(tempfile.mkdtemp(prefix="acies-workroom-e2e-"))
    project_root = fixture_root / "Nelson" / "BAC" / "2025" / "250439 E2E Project"
    saved_project_path = project_root / "Arch" / "2025-08-18 DD90"
    electrical_dir = project_root / "Electrical"
    mechanical_dir = project_root / "Mechanical"
    arch_dir = project_root / "Arch"
    saved_project_path.mkdir(parents=True, exist_ok=True)
    electrical_dir.mkdir(parents=True, exist_ok=True)
    mechanical_dir.mkdir(parents=True, exist_ok=True)
    arch_dir.mkdir(parents=True, exist_ok=True)
    (electrical_dir / "E2E-001.dwg").write_text("", encoding="utf-8")
    (electrical_dir / "E2E-002.dwg").write_text("", encoding="utf-8")
    (mechanical_dir / "M-001.dwg").write_text("", encoding="utf-8")
    (mechanical_dir / "M-002.dwg").write_text("", encoding="utf-8")
    (arch_dir / "A-001.dwg").write_text("", encoding="utf-8")

    existing_settings = read_json_or_default(SETTINGS_FILE, {})
    fixture_settings = build_fixture_settings(existing_settings)
    fixture_tasks = build_fixture_tasks(saved_project_path)

    run_result = {}

    with BackupFile(SETTINGS_FILE), BackupFile(TASKS_FILE):
        write_json(SETTINGS_FILE, fixture_settings)
        write_json(TASKS_FILE, fixture_tasks)

        api = Api()
        if not api.test_mode:
            raise RuntimeError(
                "E2E runner requires ACIES_TEST_MODE=1. Environment override failed."
            )

        window = webview.create_window(
            "ACIES E2E Workroom Validation",
            str(BASE_DIR / "index.html"),
            js_api=api,
            width=1400,
            height=900,
            resizable=True,
            min_size=(1024, 768),
        )
        webview.start(
            run_automation,
            window,
            api,
            artifacts_dir,
            project_root,
            saved_project_path,
            args.timeout,
            run_result,
        )

    if not args.keep_fixtures:
        shutil.rmtree(fixture_root, ignore_errors=True)

    report = {
        "status": run_result.get("status", "error"),
        "timestamp": datetime.utcnow().isoformat() + "Z",
        "artifacts_dir": str(artifacts_dir),
        "fixture_root": str(fixture_root),
        "result": run_result,
    }
    report_path = artifacts_dir / "report.json"
    write_json(report_path, report)

    print(f"E2E report: {report_path}")
    if report["status"] != "success":
        print(report.get("result", {}).get("message", "E2E run failed."))
        return 1

    print("E2E validation passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
