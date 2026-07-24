"""Test encoding recovery for mojibake strings in server.py"""

# Read file as raw bytes to avoid any interpretation issues
with open(r'c:\Users\hotabuch\Desktop\deepseek\server.py', 'rb') as f:
    raw = f.read()

# The file has a BOM, skip it
if raw[:3] == b'\xef\xbb\xbf':
    raw = raw[3:]

# Find the stage name strings by searching for the surrounding ASCII context
# Look for: safe_confirm_with_buttons(... , "...", 
# The stage name is in quotes after the 2nd comma argument

import re

# Find all occurrences of the pattern
# The stage names are Russian text in quotes
pattern = rb'safe_confirm_with_buttons\([^)]+\)'
matches = re.findall(pattern, raw)
print(f"Found {len(matches)} safe_confirm_with_buttons calls")

for i, m in enumerate(matches):
    text = m.decode('utf-8')
    # Extract the 3rd argument (the quoted stage name)
    parts = text.split(',')
    if len(parts) >= 3:
        arg = parts[2].strip()
        # Remove quotes
        if arg.startswith('"') and arg.endswith('"'):
            arg = arg[1:-1]
        elif arg.startswith('"'):
            # Might have more after the quote
            end = arg.find('"', 1)
            if end > 0:
                arg = arg[1:end]
        print(f"\nStage {i+1}: {repr(arg)}")
        
        # Try different recovery methods
        for enc in ['cp1251', 'latin-1', 'iso-8859-5', 'cp1252', 'koi8-r']:
            try:
                recovered = arg.encode(enc).decode('utf-8')
                if any('\u0400' <= c <= '\u04FF' for c in recovered):
                    print(f"  RECOVERED via {enc}: {recovered}")
                    break
            except:
                pass
        else:
            print(f"  Could not recover. First 10 chars: {[f'U+{ord(c):04X}' for c in arg[:10]]}")


