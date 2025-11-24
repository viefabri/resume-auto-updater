import json
import openpyxl
from openpyxl.styles import Border, Side
from openpyxl.utils import range_boundaries
from openpyxl.cell.cell import MergedCell
import os
import datetime
import copy
import sys

# Configuration
MASTER_JSON_PATH = os.path.join('005_ToolOutput', '01_ResumeUpdater', 'Data', 'resume_master.json')
UPDATE_JSON_PATH = os.path.join('005_ToolOutput', '02_ResumeUpdate', 'Data', 'resume_update.json')
TEMPLATE_EXCEL_PATH = '経歴書（gotou_ryujirou）202508.xlsx'
TEMPLATE_SHEET_NAME = '_Template'
TARGET_SHEET_NAME = 'スキルシート'
START_ROW = 21

def load_json(path):
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)

def save_json(path, data):
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

def merge_data(master_data, update_data):
    """Merges update_data into master_data."""
    print("Merging data...")
    payloads = update_data.get('update_payload', [])
    
    # Process payloads
    for payload in payloads:
        action = payload.get('action')
        target_no = payload.get('target_no')
        new_data = payload.get('data')

        if action == 'INSERT' and target_no == 0:
            master_data['work_history'].insert(0, new_data)
        elif action == 'UPDATE':
            found = False
            for i, entry in enumerate(master_data['work_history']):
                if entry.get('no') == str(target_no):
                    master_data['work_history'][i] = new_data
                    found = True
                    break
            if not found:
                print(f"Warning: Entry with no {target_no} not found for UPDATE.")

    # Renumber
    for i, entry in enumerate(master_data['work_history']):
        entry['no'] = str(i + 1)
        
    # Update Footer
    footer_update = update_data.get('footer_update')
    if footer_update and footer_update.get('update_required'):
        master_data['footer']['other_col_b'] = footer_update.get('other_col_b', [])

    return master_data

def clean_sheet(ws):
    """Clears content and removes merged cells from START_ROW onwards."""
    print(f"Cleaning sheet from row {START_ROW}...")
    
    # 1. Delete Rows
    max_row = ws.max_row
    if max_row >= START_ROW:
        ws.delete_rows(START_ROW, amount=max_row - START_ROW + 1)
        
    # 2. Remove Residual Merged Cells
    # Iterate backwards to safely remove items from list while iterating
    # or collect first then remove.
    to_remove = []
    for merged_range in ws.merged_cells.ranges:
        if merged_range.max_row >= START_ROW:
            to_remove.append(merged_range)
            
    for mr in to_remove:
        try:
            ws.merged_cells.remove(mr)
        except KeyError:
            pass # Already removed
        except Exception as e:
            print(f"Warning: Failed to remove merged range {mr}: {e}")

def copy_style(src_cell, dst_cell):
    if src_cell.has_style:
        dst_cell.font = copy.copy(src_cell.font)
        dst_cell.border = copy.copy(src_cell.border)
        dst_cell.fill = copy.copy(src_cell.fill)
        dst_cell.number_format = copy.copy(src_cell.number_format)
        dst_cell.protection = copy.copy(src_cell.protection)
        dst_cell.alignment = copy.copy(src_cell.alignment)

def safe_write(ws, row, col, value):
    """Writes value to cell ONLY if it is NOT a MergedCell."""
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        # Skip writing to merged cells (read-only)
        return
    cell.value = value

def apply_template_and_write_data(ws, template_ws, entry, start_row):
    """Applies template and writes data for one entry (5 rows)."""
    
    # 1. Copy Template (Styles & Merges)
    for row_idx in range(1, 6):
        for col_idx in range(1, 32): # A to AE
            src_cell = template_ws.cell(row=row_idx, column=col_idx)
            dst_cell = ws.cell(row=start_row + row_idx - 1, column=col_idx)
            copy_style(src_cell, dst_cell)

    # Copy Merges
    for merged_range in template_ws.merged_cells.ranges:
        min_col, min_row, max_col, max_row = range_boundaries(str(merged_range))
        new_min_row = min_row + start_row - 1
        new_max_row = max_row + start_row - 1
        ws.merge_cells(start_row=new_min_row, start_column=min_col, end_row=new_max_row, end_column=max_col)

    # 2. Write Data using safe_write
    
    # No (A)
    safe_write(ws, start_row, 1, entry.get('no'))
    
    # Period (B)
    safe_write(ws, start_row, 2, entry.get('period', {}).get('start'))
    safe_write(ws, start_row + 2, 2, entry.get('period', {}).get('end'))
    
    # Business Content
    bc = entry.get('business_content', {})
    titles = bc.get('title_col_e', [])
    roles = bc.get('role_col_f', [])
    details = bc.get('detail_col_g', [])
    
    # Ensure lists have 5 elements (padding)
    titles += [''] * (5 - len(titles))
    roles += [''] * (5 - len(roles))
    details += [''] * (5 - len(details))
    
    for i in range(5):
        safe_write(ws, start_row + i, 5, titles[i])
        safe_write(ws, start_row + i, 6, roles[i])
        safe_write(ws, start_row + i, 7, details[i])

    # Technology
    tech = entry.get('technology', {})
    envs = tech.get('environment_col_u', [])
    langs = tech.get('language_col_z', [])
    procs = tech.get('process_col_ae', [])
    
    envs += [''] * (5 - len(envs))
    langs += [''] * (5 - len(langs))
    procs += [''] * (5 - len(procs))
    
    for i in range(5):
        safe_write(ws, start_row + i, 21, envs[i])
        safe_write(ws, start_row + i, 26, langs[i])
        safe_write(ws, start_row + i, 31, procs[i])

def write_footer(ws, footer_data, start_row):
    print("Writing footer...")
    others = footer_data.get('other_col_b', [])
    
    safe_write(ws, start_row, 1, "その他")
    
    for i, line in enumerate(others):
        safe_write(ws, start_row + i, 2, line)

def draw_border(ws, start_row, end_row):
    print("Drawing borders...")
    medium = Side(border_style="medium", color="000000")
    
    # Helper to get style-able cell (top-left if merged)
    def get_style_cell(r, c):
        cell = ws.cell(row=r, column=c)
        if isinstance(cell, MergedCell):
            for mr in ws.merged_cells.ranges:
                if cell.coordinate in mr:
                    return ws.cell(row=mr.min_row, column=mr.min_col)
        return cell

    # Top & Bottom
    for col in range(1, 32):
        # Top
        cell = get_style_cell(start_row, col)
        new_border = copy.copy(cell.border)
        new_border.top = medium
        cell.border = new_border
        
        # Bottom
        cell = get_style_cell(end_row, col)
        new_border = copy.copy(cell.border)
        new_border.bottom = medium
        cell.border = new_border

    # Left & Right
    for row in range(start_row, end_row + 1):
        # Left (A)
        cell = get_style_cell(row, 1)
        new_border = copy.copy(cell.border)
        new_border.left = medium
        cell.border = new_border
        
        # Right (AE)
        cell = get_style_cell(row, 31)
        new_border = copy.copy(cell.border)
        new_border.right = medium
        cell.border = new_border

def main():
    # 1. Load & Merge
    try:
        master_data = load_json(MASTER_JSON_PATH)
        update_data = load_json(UPDATE_JSON_PATH)
    except FileNotFoundError as e:
        print(f"Error: {e}")
        return

    master_data = merge_data(master_data, update_data)
    save_json(MASTER_JSON_PATH, master_data)
    print(f"Updated {MASTER_JSON_PATH}")

    # 2. Open Excel
    try:
        wb = openpyxl.load_workbook(TEMPLATE_EXCEL_PATH)
    except FileNotFoundError:
        print(f"Error: Excel template not found: {TEMPLATE_EXCEL_PATH}")
        return

    if TARGET_SHEET_NAME not in wb.sheetnames or TEMPLATE_SHEET_NAME not in wb.sheetnames:
        print(f"Error: Missing sheets. Required: {TARGET_SHEET_NAME}, {TEMPLATE_SHEET_NAME}")
        print(f"Available sheets: {wb.sheetnames}")
        return

    ws = wb[TARGET_SHEET_NAME]
    template_ws = wb[TEMPLATE_SHEET_NAME]

    # 3. Clean
    clean_sheet(ws)

    # 4. Render
    current_row = START_ROW
    print("Rendering history...")
    for entry in master_data['work_history']:
        apply_template_and_write_data(ws, template_ws, entry, current_row)
        current_row += 5

    # 5. Footer
    write_footer(ws, master_data['footer'], current_row)
    
    # 6. Border
    # Footer length
    footer_len = len(master_data['footer'].get('other_col_b', []))
    # Border covers history + footer? Usually just history in this context, 
    # but let's assume history area based on previous context.
    # Actually, looking at the template, the border usually wraps the whole content.
    # Let's wrap up to the last row written.
    last_row = current_row + max(0, footer_len - 1)
    if footer_len == 0: last_row = current_row - 1 # Just history
    
    # Wait, the user requirement said "业务履歴エリア全体 (No.1〜最終行)".
    # This implies the border should enclose the work history blocks.
    # Footer usually has its own style or is outside.
    # Let's stick to history area: START_ROW to current_row - 1.
    draw_border(ws, START_ROW, current_row - 1)

    # 7. Save
    timestamp = datetime.datetime.now().strftime('%Y%m%d')
    output_filename = f"経歴書_Updated_{timestamp}.xlsx"
    wb.save(output_filename)
    print(f"Success! Saved to {output_filename}")

if __name__ == "__main__":
    main()
