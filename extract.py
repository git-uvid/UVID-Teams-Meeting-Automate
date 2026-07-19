import re
content = open('Ref Files/timezone.txt', 'r', encoding='utf-8').read()
matches = re.findall(r'<span data-automation-id=\"flow-combobox-option\">(.*?)</span>', content)
open('temp_tz.html', 'w', encoding='utf-8').write('\n'.join([f'            <option value=\"{m}\">{m}</option>' for m in matches]))
