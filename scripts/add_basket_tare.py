"""Add a blank tare-weight column to INFO without rewriting Excel extensions."""
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
import os,re,tempfile
path=Path('Mazandasht_Tamplate V00/INFO.xlsx')
ns={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
with ZipFile(path) as src:
 wb=ET.fromstring(src.read('xl/workbook.xml'));rels=ET.fromstring(src.read('xl/_rels/workbook.xml.rels'))
 targets={r.attrib['Id']:r.attrib['Target'].lstrip('/') for r in rels}
 sh=next(x for x in wb.find('m:sheets',ns) if x.attrib['name']=='Baskets')
 target=targets[sh.attrib['{'+ns['r']+'}id']];target=target if target.startswith('xl/') else 'xl/'+target
 fd,tmp=tempfile.mkstemp(dir=path.parent,suffix='.xlsx');os.close(fd)
 with ZipFile(tmp,'w') as out:
  for item in src.infolist():
   data=src.read(item.filename)
   if item.filename==target:
    text=data.decode('utf-8')
    if re.search(r'<c\s+[^>]*\br="F5"',text):raise ValueError('F5 already exists')
    def row(m):
     r=int(m.group(1)); xml=m.group(0)
     if r==5: cell='<c r="F5" s="3" t="inlineStr"><is><t>وزن تار (کیلوگرم)</t></is></c>'
     elif r==6:cell='<c r="F6" s="4" t="inlineStr"><is><t>Tare weight (kg)</t></is></c>'
     elif 7<=r<=36:cell=f'<c r="F{r}" s="{14 if r%2 else 15}"/>'
     else:return xml
     return xml.replace('</row>',cell+'</row>')
    text=re.sub(r'<row\s[^>]*\br="(\d+)"[^>]*>.*?</row>',row,text,flags=re.S)
    text=re.sub(r'(<dimension ref=")([^"]+)("\s*/>)',lambda m:m.group(1)+'A1:F36'+m.group(3),text,count=1)
    data=text.encode()
   out.writestr(item,data)
os.replace(tmp,path)
