import zipfile, xml.etree.ElementTree as ET
import pandas as pd
from pathlib import Path

base = Path(r'c:\Users\matukunda\OneDrive\Documents\FIRE_ENGINEERING')
print('BASE:', base)
path = base / 'Fire_Engineering_Risk_Profile_Data_Dictionary_v2_0.docx'
print('DOCX exists:', path.exists())
with zipfile.ZipFile(path) as z:
    root = ET.fromstring(z.read('word/document.xml'))
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    paras = []
    for p in root.findall('.//w:p', ns):
        texts = [t.text or '' for t in p.findall('.//w:t', ns)]
        txt = ''.join(texts).strip()
        if txt:
            paras.append(txt)
    print('DOCX PARAGRAPHS:')
    for i, p in enumerate(paras[:120], 1):
        print(f'{i}: {p[:500]}')

print('\n---\n')
xl = pd.ExcelFile(base / 'data.xlsx')
print('SHEETS:', xl.sheet_names)
for sheet in xl.sheet_names:
    df = xl.parse(sheet)
    print(f'\nSHEET: {sheet}')
    print('shape:', df.shape)
    print(df.head(8).to_string(index=False))
    print('columns:', list(df.columns[:20]))
    print('dtypes:', df.dtypes.astype(str).to_dict())
