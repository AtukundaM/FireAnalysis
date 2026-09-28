from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import seaborn as sns

base = Path(r'c:\Users\matukunda\OneDrive\Documents\FIRE_ENGINEERING')

def read_docx_text(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read('word/document.xml'))
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    paragraphs = []
    for p in root.findall('.//w:p', ns):
        texts = [t.text or '' for t in p.findall('.//w:t', ns)]
        text = ''.join(texts).strip()
        if text:
            paragraphs.append(text)
    return paragraphs

# Document text
print('=== WORD DOCUMENT TEXT ===')
docx_path = base / 'Fire_Engineering_Risk_Profile_Data_Dictionary_v2_0.docx'
paragraphs = read_docx_text(docx_path)
for i, p in enumerate(paragraphs[:200], 1):
    print(f'{i}: {p[:600]}')

print('\n=== EXCEL SHEETS ===')
xl_path = base / 'data.xlsx'
xl = pd.ExcelFile(xl_path)
print('Sheets:', xl.sheet_names)
for sheet in xl.sheet_names:
    df = pd.read_excel(xl_path, sheet_name=sheet)
    print(f'\n--- {sheet} ---')
    print('shape:', df.shape)
    print(df.head(10).to_string(index=False))
    print('columns:', list(df.columns[:20]))
    print('dtypes:', df.dtypes.astype(str).to_dict())
