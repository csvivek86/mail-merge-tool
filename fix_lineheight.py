#!/usr/bin/env python3

# Fix the LineHeightTypes issue in app.py
with open('src/app.py', 'r') as f:
    content = f.read()

# Replace the problematic line
old_text = 'QTextBlockFormat.LineHeightTypes.ProportionalHeight'
new_text = '1'

fixed_content = content.replace(old_text, new_text)

with open('src/app.py', 'w') as f:
    f.write(fixed_content)

print(f"Replaced {content.count(old_text)} occurrences of {old_text} with {new_text}")
