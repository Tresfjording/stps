from pathlib import Path
path = Path(r'c:\Users\ØyvindGranberg\Projects\stps\blabla.py')
text = path.read_text(encoding='utf-8')
start = text.index('def build_player_history_section_html(history_path: Path) -> str:')
end = text.index('def export_html_to_pdf', start)
func = text[start:end]
if 'return """' in func:
    body_start = func.index('return """') + len('return """')
    body_end = func.rindex('""" % (options_html, player_data_json)')
    body = func[body_start:body_end]
    body = body.replace('{{', '{').replace('}}', '}')
    new_func = func[:body_start] + body + func[body_end:]
    text = text[:start] + new_func + text[end:]
    path.write_text(text, encoding='utf-8')
    print('rewrote chart template braces')
else:
    raise SystemExit('did not find function body')
