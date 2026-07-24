"""Fix double-encoded UTF-8 in server.py"""
import os

path = r'c:\Users\hotabuch\Desktop\deepseek\server.py'

# Read the file, stripping BOM if present
with open(path, 'r', encoding='utf-8-sig') as f:
    content = f.read()

# Reverse double-encoding: Latin-1 -> UTF-8
# The BOM was stripped by utf-8-sig, now all chars should be Latin-1 encodable
fixed = content.encode('latin-1').decode('utf-8')

# Verify fix worked by checking common Cyrillic words
checks = ['\u042d\u0442\u0430\u043f', '\u0424\u0418\u041e', '\u041f\u0440\u043e\u0431\u0438\u0432', '\u041d\u043e\u043c\u0435\u0440']
all_ok = all(c in fixed for c in checks)

if all_ok:
    # Backup original
    bak = path + '.bak'
    with open(bak, 'w', encoding='utf-8') as f:
        f.write(content)
    print(f'Backup saved: {bak}')
    # Write fixed without BOM (standard UTF-8)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(fixed)
    print('SUCCESS: file re-encoded to proper UTF-8 Cyrillic!')
else:
    print('WARNING: fixed file missing expected Cyrillic patterns.')
    for c in checks:
        print(f'  Check "{c}": {"OK" if c in fixed else "MISSING"}')
    print('Aborting without changes.')
