#!/usr/bin/env python3
# Fix LineHeightTypes issue in app.py

with open('src/app.py', 'r') as f:
    lines = f.readlines()

# Fix line 271 (index 270)
if len(lines) > 270:
    lines[270] = lines[270].replace('QTextBlockFormat.LineHeightTypes.ProportionalHeight', '1')
    print(f"Fixed line 271: {lines[270].strip()}")

# Fix line 666 (index 665)
if len(lines) > 665:
    lines[665] = lines[665].replace('QTextBlockFormat.LineHeightTypes.ProportionalHeight', '1')
    print(f"Fixed line 666: {lines[665].strip()}")

with open('src/app.py', 'w') as f:
    f.writelines(lines)

print("LineHeightTypes fix completed!")
