"""Clear explicitly marked sample input cells without rewriting Excel features."""
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
import os, tempfile
NS={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships','p':'http://schemas.openxmlformats.org/package/2006/relationships'}
base=Path('Mazandasht_Tamplate V00')
plans={'Transactions.xlsx':{'Entry':list(range(9,31))},'Cheques.xlsx':{'Issues':list(range(9,13))},'Employees.xlsx':{'Wages':list(range(9,13))},'INFO.xlsx':{'Accounts':[7,8,9,210,411,412,507,508,509,537,557],'Staff':[11,12]}}
for name,sheets in plans.items():
 path=base/name
 with ZipFile(path) as zin:
  wb=ET.fromstring(zin.read('xl/workbook.xml')); rel=ET.fromstring(zin.read('xl/_rels/workbook.xml.rels'))
  targets={x.attrib['Id']:x.attrib['Target'].lstrip('/') for x in rel}
  edits={}
  for sh in wb.find('m:sheets',NS):
   if sh.attrib['name'] in sheets:
    target=targets[sh.attrib['{'+NS['r']+'}id']]; target=target if target.startswith('xl/') else 'xl/'+target
    edits[target]=set(sheets[sh.attrib['name']])
  fd,tmp=tempfile.mkstemp(dir=path.parent,suffix='.xlsx');os.close(fd)
  with ZipFile(tmp,'w') as out:
   for item in zin.infolist():
    data=zin.read(item.filename)
    if item.filename in edits:
     # Replace individual XML cells only, preserving namespaces and workbook extensions.
     import re
     rows=edits[item.filename]
     def clear(m):
      cell=m.group(0); ref=re.search(rb'\br="([A-Z]+)(\d+)"',cell)
      if cell.endswith(b'/>' ) or not ref or int(ref.group(2)) not in rows or re.search(rb'<(?:\w+:)?f(?:\s|>)',cell):return cell
      # retain style, remove type and contents
      start=cell[:cell.index(b'>')];start=re.sub(rb'\s+t="[^"]*"',b'',start)
      return start+b'/>'
     data=re.sub(rb'<c\s[^>]*?(?:/>|>.*?</c>)',clear,data,flags=re.S)
    out.writestr(item,data)
 os.replace(tmp,path)
