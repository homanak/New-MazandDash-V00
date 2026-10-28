"""Repair the basket calculation engine's shifted Entry references in-place.
Keeps Excel-only extensions and external links intact. Run on the original file once.
"""
from pathlib import Path
from zipfile import ZipFile
import os,re,tempfile
from xml.etree import ElementTree as ET
p=Path('Mazandasht_Tamplate V00/Baskets.xlsx')
ns={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
with ZipFile(p) as src:
 w=ET.fromstring(src.read('xl/workbook.xml')); rel=ET.fromstring(src.read('xl/_rels/workbook.xml.rels'))
 ids={x.attrib['Id']:x.attrib['Target'].lstrip('/') for x in rel}
 sh=next(x for x in w.find('m:sheets',ns) if x.attrib['name']=='Calc')
 target=ids[sh.attrib['{'+ns['r']+'}id']]; target=target if target.startswith('xl/') else 'xl/'+target
 fd,tmp=tempfile.mkstemp(dir=p.parent,suffix='.xlsx');os.close(fd)
 with ZipFile(tmp,'w') as dst:
  for info in src.infolist():
   data=src.read(info.filename)
   if info.filename==target:
    text=data.decode('utf-8');count=[0]
    def fix(m):
     col=m.group(1); row=int(m.group(2));
     if row<6 or row>1005:return m.group(0)
     targetrow=row+3; ref=f'Entry!${col}{targetrow}'
     formula=f'IF($A{row}="","",'+(ref if col=='G' else f'IF({ref}="",0,{ref})')+')'
     count[0]+=1
     return re.sub(r'<f(?:\s[^>]*)?>.*?</f>\s*(?:<v(?:\s[^>]*)?>.*?</v>|<v\s*/>)?', '<f>'+formula.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')+'</f><v></v>',m.group(0),flags=re.S)
    text=re.sub(r'<c\s+[^>]*\br="([GHI])(\d+)"[^>]*>.*?</c>',fix,text,flags=re.S)
    assert count[0]==3000,count
    data=text.encode('utf-8')
   elif info.filename=='xl/workbook.xml':
    text=data.decode('utf-8');text=re.sub(r'<calcPr\b[^>]*/>', '<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>',text);data=text.encode('utf-8')
   dst.writestr(info,data)
os.replace(tmp,p)
