import json
import os
import copy
import sys

# Paths
MASTER_JSON_PATH = os.path.join('005_ToolOutput', '01_ResumeUpdater', 'Data', 'resume_master.json')
DRAFT_JSON_PATH = os.path.join('005_ToolOutput', '02_ResumeUpdate', 'Data', 'resume_update.json')
OUTPUT_DIR = os.path.join('005_ToolOutput', '03_PlanResult')
OUTPUT_FILE = os.path.join(OUTPUT_DIR, 'resume_merged_preview.json')

def load_json(path):
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"Error: File not found at {path}")
        sys.exit(1)
    except json.JSONDecodeError:
        print(f"Error: Invalid JSON at {path}")
        sys.exit(1)

def save_json(path, data):
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

def validate_payload(payloads):
    """Validates and auto-corrects the payload structure."""
    print("Validation: Checking payload structure...")
    fixed_count = 0
    
    list_fields = [
        ('business_content', 'title_col_e'),
        ('business_content', 'role_col_f'),
        ('business_content', 'detail_col_g'),
        ('technology', 'environment_col_u'),
        ('technology', 'language_col_z'),
        ('technology', 'process_col_ae')
    ]

    for idx, item in enumerate(payloads):
        # Check required fields
        if 'action' not in item or 'target_no' not in item or 'data' not in item:
            print(f"  [Error] Item {idx}: Missing action, target_no, or data.")
            continue
            
        data = item['data']
        
        # Check list lengths
        for parent, field in list_fields:
            if parent in data and field in data[parent]:
                current_list = data[parent][field]
                if not isinstance(current_list, list):
                    print(f"  [Warning] Item {idx}: {field} is not a list. Converting to list.")
                    current_list = [str(current_list)]
                    data[parent][field] = current_list
                    fixed_count += 1
                
                if len(current_list) != 5:
                    print(f"  [Fix] Item {idx}: {field} length is {len(current_list)}. Padding/Truncating to 5.")
                    # Pad with empty strings
                    while len(current_list) < 5:
                        current_list.append("")
                    # Truncate if too long (though padding is the main requirement)
                    if len(current_list) > 5:
                        current_list = current_list[:5]
                    
                    data[parent][field] = current_list
                    fixed_count += 1
    
    if fixed_count > 0:
        print(f"Validation: Fixed {fixed_count} issues.")
    else:
        print("Validation: OK")
    
    return payloads

def simulate_merge(master_data, draft_data):
    """Simulates the merge process."""
    print("\nSimulation: Starting merge simulation...")
    
    merged_data = copy.deepcopy(master_data)
    payloads = draft_data.get('update_payload', [])
    
    # Validate and Fix
    payloads = validate_payload(payloads)
    
    initial_count = len(merged_data['work_history'])
    
    for item in payloads:
        action = item.get('action')
        target_no = item.get('target_no')
        new_data = item.get('data')
        
        print(f"  Action Detected: {action} (Target: {target_no})")
        
        if action == 'INSERT' and target_no == 0:
            merged_data['work_history'].insert(0, new_data)
        elif action == 'UPDATE':
            # Find by 'no'
            found = False
            for i, entry in enumerate(merged_data['work_history']):
                if entry.get('no') == str(target_no):
                    merged_data['work_history'][i] = new_data
                    found = True
                    break
            if not found:
                print(f"  [Warning] Update target No.{target_no} not found.")
        else:
            print(f"  [Warning] Unknown action: {action}")

    # Renumbering
    print("  Renumbering entries...")
    for i, entry in enumerate(merged_data['work_history']):
        entry['no'] = str(i + 1)
        
    final_count = len(merged_data['work_history'])
    print(f"Impact: Total entries {initial_count} -> {final_count}")
    
    # Footer Update
    footer_update = draft_data.get('footer_update')
    if footer_update and footer_update.get('update_required'):
        print("  Footer Update: Applied.")
        merged_data['footer']['other_col_b'] = footer_update.get('other_col_b', [])
    else:
        print("  Footer Update: None.")
        
    return merged_data

def main():
    # Ensure output directory exists
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    # Load Data
    print(f"Loading Master: {MASTER_JSON_PATH}")
    master_data = load_json(MASTER_JSON_PATH)
    
    print(f"Loading Draft: {DRAFT_JSON_PATH}")
    draft_data = load_json(DRAFT_JSON_PATH)
    
    # Simulate
    merged_data = simulate_merge(master_data, draft_data)
    
    # Save
    save_json(OUTPUT_FILE, merged_data)
    print(f"\nSaved: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
