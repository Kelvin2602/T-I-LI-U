import os
import docx

base_dir = r'd:\B\New folder\T-I-LI-U'

for filename in os.listdir(base_dir):
    if filename.endswith('.docx') and not filename.startswith('~'):
        filepath = os.path.join(base_dir, filename)
        try:
            doc = docx.Document(filepath)
            text_lines = []
            for para in doc.paragraphs:
                if para.text.strip():
                    text_lines.append(para.text.strip())
            content = '\n'.join(text_lines)
            print(f'\n{"="*80}')
            print(f'FILE: {filename}')
            print(f'{"="*80}')
            print(content[:5000])
        except Exception as e:
            print(f'\nFILE: {filename} - ERROR: {e}')
