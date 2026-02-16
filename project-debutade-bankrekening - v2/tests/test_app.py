import os
import json
import tempfile
import shutil
from datetime import datetime

import openpyxl


def create_temp_workbook(xlsx_path, required_sheets, required_headers):
    wb = openpyxl.Workbook()
    # Ensure first sheet name matches first required
    first = required_sheets[0]
    ws = wb.active
    ws.title = first
    # Write headers
    ws.append(required_headers)
    # Add a few rows: two untagged, one tagged
    ws.append([datetime(2026, 1, 1), 'Omschrijving A', 'ING', 'NL00', 'CODE', 'Af', 10.0, 'Type', 'Memo A', 100.0, '', ''])
    ws.append([datetime(2026, 1, 2), 'Omschrijving B', 'ING', 'NL00', 'CODE', 'Bij', 20.5, 'Type', 'Memo B', 120.5, '', ''])
    ws.append([datetime(2026, 1, 3), 'Omschrijving C', 'ING', 'NL00', 'CODE', 'Bij', 5.0, 'Type', 'Memo C', 125.5, '', '500;Vermogen Debutade'])

    # Other sheets
    for name in required_sheets[1:]:
        s = wb.create_sheet(title=name)
        s.append(required_headers)
        s.append([datetime(2026, 2, 1), 'Spaar A', 'ING Spaar', 'NL00', 'CODE', 'Bij', 1.0, 'Type', 'Memo S', 1000.0, '', ''])
    wb.save(xlsx_path)
    wb.close()


def create_temp_config(config_path, xlsx_path, required_sheets, tags):
    base_dir = os.path.dirname(config_path)
    backup_dir = os.path.join(base_dir, 'backup')
    log_dir = os.path.join(base_dir, 'log')
    os.makedirs(backup_dir, exist_ok=True)
    os.makedirs(log_dir, exist_ok=True)
    data = {
        "excel_file_path": xlsx_path,
        "backup_directory": backup_dir,
        "resources": base_dir,
        "log_directory": log_dir,
        "excel_sheet_name": required_sheets[0],
        "required_sheets": required_sheets,
        "tags": tags,
        "log_level": "INFO"
    }
    with open(config_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=4)
    return data


def run_tests():
    # Prepare temp workspace under project tests folder
    tmp_root = tempfile.mkdtemp(prefix='bankrekening_tests_')
    # Dynamic import via spec to avoid path/package issues
    import importlib.util as util
    spec = util.spec_from_file_location("webapp", os.path.join("c:\\project-debutade-bankrekening", "webapp.py"))
    webapp = util.module_from_spec(spec)
    spec.loader.exec_module(webapp)

    # Build temp workbook + config
    xlsx_path = os.path.join(tmp_root, 'test.xlsx')
    create_temp_workbook(xlsx_path, webapp.REQUIRED_SHEETS, webapp.REQUIRED_HEADERS)

    temp_config = os.path.join(tmp_root, 'config.json')
    create_temp_config(temp_config, xlsx_path, webapp.REQUIRED_SHEETS, webapp.TAGS)

    # Point app to temp config and re-import webapp to pick it up
    os.environ['BANKREKENING_CONFIG'] = temp_config
    import importlib.util as util
    spec = util.spec_from_file_location("webapp", os.path.join("c:\\project-debutade-bankrekening", "webapp.py"))
    webapp = util.module_from_spec(spec)
    spec.loader.exec_module(webapp)

    print("== Sheet stats before ==")
    stats_before = webapp.get_sheet_stats()
    print(stats_before)

    print("== Untagged before ==")
    untagged_before = webapp.get_untagged_transactions()
    print(len(untagged_before), "rows")
    # Choose first untagged row
    if untagged_before:
        target = untagged_before[0]
        sheet = target['sheet_name']
        row = target['row_index']
        # Use Flask test client to hit update_tag
        client = webapp.app.test_client()
        resp = client.post('/update_tag', json={
            'sheet_name': sheet,
            'row_index': row,
            'tag': webapp.TAGS[0] if webapp.TAGS else 'TestTag'
        })
        print("update_tag status:", resp.status_code, resp.json)
    else:
        print("No untagged rows found to update.")

    print("== Sheet stats after ==")
    stats_after = webapp.get_sheet_stats()
    print(stats_after)

    print("== Untagged after ==")
    untagged_after = webapp.get_untagged_transactions()
    print(len(untagged_after), "rows")

    # Render index and show section presence
    client = webapp.app.test_client()
    index_resp = client.get('/')
    print("index status:", index_resp.status_code)
    html = index_resp.data.decode('utf-8')
    print("Contains 'Transacties zonder Tag':", "Transacties zonder Tag" in html)
    print("Contains 'Alle transacties':", "Alle transacties" in html)

    # Validate workbook structure
    ok, msg = webapp.validate_workbook_structure(webapp.EXCEL_FILE_PATH)
    print("Workbook structure valid:", ok, "message:", msg)

    return {
        'stats_before': stats_before,
        'stats_after': stats_after,
        'untagged_before': len(untagged_before),
        'untagged_after': len(untagged_after)
    }


if __name__ == '__main__':
    results = run_tests()
    print("== Summary ==")
    print(results)
