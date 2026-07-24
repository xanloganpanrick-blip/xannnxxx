"""Fix stage name strings in server.py by replacing 3rd arg of safe_confirm_with_buttons calls."""
import re

# The correct stage names as Unicode escape sequences
CORRECT_STAGES = [
    "\u042d\u0442\u0430\u043f 1: \u0424\u0418\u041e+\u041d\u041e\u041c\u0415\u0420",
    "\u042d\u0442\u0430\u043f 2: \u0421\u041d\u0418\u041b\u0421 (\u0431\u043e\u04421)",
    "\u042d\u0442\u0430\u043f 3: \u0421\u041d\u0418\u041b\u0421 \u0434\u043e\u0431\u0438\u0432\u043a\u0430 (\u0431\u043e\u04422)",
    "\u042d\u0442\u0430\u043f 3.5: \u041f\u0420\u041e\u0411\u0418\u0412 \u041f\u041e \u041d\u041e\u041c\u0415\u0420\u0423",
    "\u042d\u0442\u0430\u043f 5: \u0424\u0418\u041e+\u0414\u0410\u0422\u0410 (\u0431\u043e\u04421)",
    "\u042d\u0442\u0430\u043f 6: \u0424\u0418\u041e+\u0414\u0410\u0422\u0410 \u0434\u043e\u0431\u0438\u0432\u043a\u0430 (\u0431\u043e\u04422)",
    "\u042d\u0442\u0430\u043f 7: \u0414\u041e\u0411\u0418\u0412 \u0421\u0410\u0423\u0420\u041e\u041d",
    "\u042d\u0442\u0430\u043f 8: \u0414\u041e\u0411\u0418\u0412 \u041a\u0412\u0410\u0420\u0422\u0418\u0420",
]

path = r'c:\Users\hotabuch\Desktop\deepseek\server.py'

# Read as bytes
with open(path, 'rb') as f:
    data = f.read()

# Skip BOM if present
if data[:3] == b'\xef\xbb\xbf':
    bom = data[:3]
    data = data[3:]
else:
    bom = b''

text = data.decode('utf-8')
lines = text.split('\n')

# Find and fix safe_confirm_with_buttons calls (not the def line)
stage_idx = 0
fixed_lines = []
for line in lines:
    if 'safe_confirm_with_buttons(' in line and 'def safe_confirm_with_buttons' not in line:
        # Replace the 3rd argument (quoted stage name)
        # Pattern: safe_confirm_with_buttons(..., ..., "STAGE_NAME", ...)
        # Find the 3rd comma-separated argument
        # Split only on commas that are not inside parentheses (simplified)
        m = re.search(r'safe_confirm_with_buttons\(([^,]+),\s*([^,]+),\s*"([^"]*)"', line)
        if m and stage_idx < len(CORRECT_STAGES):
            old_stage = m.group(3)
            new_stage = CORRECT_STAGES[stage_idx]
            line = line.replace(f'"{old_stage}"', f'"{new_stage}"', 1)
            print(f"Fixed stage {stage_idx+1}: {repr(old_stage)[:50]}... -> {new_stage}")
            stage_idx += 1
    fixed_lines.append(line)

if stage_idx > 0:
    # Backup
    bak = path + '.bak2'
    with open(bak, 'wb') as f:
        f.write(bom + data)
    print(f'Backup saved: {bak}')
    # Write fixed
    result = bom + '\n'.join(fixed_lines).encode('utf-8')
    with open(path, 'wb') as f:
        f.write(result)
    print(f'SUCCESS: {stage_idx} stage names fixed!')
else:
    print('ERROR: No stage names were fixed. Check the regex pattern.')
