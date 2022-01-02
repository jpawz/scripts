#! python3
# clear_copied_text.py - Removes redundant whitespaces and new lines for text 
#                        in clipboard (e.x. copied from PDF)

import pyperclip, re

text = str(pyperclip.paste())

text = ' '.join(text.split())
text = re.sub(r'\s([?.!"](?:\s|$))', r'\1', text)

pyperclip.copy(text)
