import re
content = open('src/taskpane/taskpane.html', 'r', encoding='utf-8').read()
options = open('temp_tz.html', 'r', encoding='utf-8').read()
new_content = re.sub(r'(<select id=\"timezone\">).*?(</select>)', r'\g<1>\n' + options + r'\n          \g<2>', content, flags=re.DOTALL)
open('src/taskpane/taskpane.html', 'w', encoding='utf-8').write(new_content)
