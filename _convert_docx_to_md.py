import os
import docx

base_dir = r'd:\B\New folder\T-I-LI-U'

converted = 0
errors = []

for filename in os.listdir(base_dir):
    if filename.endswith('.docx') and not filename.startswith('~'):
        filepath = os.path.join(base_dir, filename)
        md_filename = filename.replace('.docx', '.md')
        md_filepath = os.path.join(base_dir, md_filename)

        if os.path.exists(md_filepath):
            print(f'SKIP (đã tồn tại): {md_filename}')
            continue

        try:
            doc = docx.Document(filepath)
            text_lines = []

            for para in doc.paragraphs:
                text = para.text.strip()
                if text:
                    style = para.style.name if para.style else ''
                    if style and 'Heading' in style:
                        level = style.replace('Heading ', '').replace('Heading', '')
                        try:
                            lv = int(level)
                            text_lines.append('#' * lv + ' ' + text)
                        except:
                            text_lines.append('# ' + text)
                    else:
                        text_lines.append(text)

            content = '\n\n'.join(text_lines)

            with open(md_filepath, 'w', encoding='utf-8') as f:
                f.write(content)

            print(f'OK: {filename} -> {md_filename}')
            converted += 1

        except Exception as e:
            errors.append(f'{filename}: {e}')
            print(f'ERROR: {filename} - {e}')

print(f'\nDone. Converted: {converted}, Errors: {len(errors)}')
if errors:
    for e in errors:
        print(f'  {e}')