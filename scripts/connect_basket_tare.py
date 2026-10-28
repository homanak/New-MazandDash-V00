"""Connect INFO basket tare to Baskol and Purchases while preserving OOXML extensions.
Run once against the unmodified workbook set. Existing dated rows get a legacy marker.
"""
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape
import re, os, tempfile
import openpyxl
BASE=Path('Mazandasht_Tamplate V00')
NS={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}

def cell(ref, formula=None, text=None, style=None):
 s=f' s="{style}"' if style is not None else ''
 if text is not None:return f'<c r="{ref}"{s} t="inlineStr"><is><t>{escape(text)}</t></is></c>'
 if formula is not None:return f'<c r="{ref}"{s} t="str"><f>{escape(formula)}</f><v></v></c>'
 return f'<c r="{ref}"{s}/>'

def rows(xml, callback):
 return re.sub(r'<row\b[^>]*\br="(\d+)"[^>]*(?:/>|>.*?</row>)',lambda m: callback(int(m.group(1)),m.group(0)),xml,flags=re.S)

def append(row, content):
 assert row.endswith('</row>')
 return row[:-6]+content+'</row>'

def replace_cell(row,ref,new):
 pattern=rf'<c\s+[^>]*\br="{ref}"[^>]*(?:/>|>.*?</c>)'
 row,n=re.subn(pattern,lambda m:new,row,count=1,flags=re.S)
 assert n==1,ref
 return row

def writebook(path, edit):
 with ZipFile(path) as src:
  book=ET.fromstring(src.read('xl/workbook.xml'));rel=ET.fromstring(src.read('xl/_rels/workbook.xml.rels'))
  paths={v.attrib['Id']:v.attrib['Target'].lstrip('/') for v in rel};sheets={}
  for sh in book.find('m:sheets',NS):
   t=paths[sh.attrib['{'+NS['r']+'}id']];sheets[sh.attrib['name']]=t if t.startswith('xl/') else 'xl/'+t
  fd,tmp=tempfile.mkstemp(dir=path.parent,suffix='.xlsx');os.close(fd)
  try:
   with ZipFile(tmp,'w') as out:
    for item in src.infolist():
     data=src.read(item.filename)
     if item.filename in edit:
      data=edit[item.filename](data.decode()).encode()
     elif item.filename in sheets.values():
      sheet=next(k for k,v in sheets.items() if v==item.filename)
      if sheet in edit:data=edit[sheet](data.decode()).encode()
     out.writestr(item,data)
   os.replace(tmp,path)
  finally:
   if os.path.exists(tmp):os.unlink(tmp)

def recalc(xml):
 return re.sub(r'<calcPr\b[^>]*/>','<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>',xml)

p=BASE/'Baskol.xlsx'
w=openpyxl.load_workbook(p,read_only=True,data_only=True)
old={r[1].row for r in w['Baskol'].iter_rows(min_row=9,max_row=2007) if r[1].value is not None};w.close()

def baskol(xml):
 assert 'r="P7"' not in xml
 def process(r,row):
  if r==4:
   # The original directions describe only P01; replace the inline content without touching style.
   return replace_cell(row,'A4',cell('A4',text='نوع سبد را در ستون P انتخاب کنید و وزن تار را در INFO ← Baskets ستون F وارد کنید. ردیف‌های قدیمی بدون نوع سبد از P01 استفاده می‌کنند؛ برای ردیف جدید با تعداد سبد، انتخاب نوع و ثبت وزن الزامی است.',style=12))
  if r==7:return append(row,cell('P7',text='نوع سبد',style=1)+cell('Q7',text='کنترل وزن سبد',style=1)+cell('R7',text='قدیمی',style=1))
  if r==8:return append(row,cell('P8',text='Basket type',style=2)+cell('Q8',text='Tare status',style=2)+cell('R8',text='Legacy',style=2))
  if 9<=r<=2007:
   # Do not silently fall back to 1.5 for a newly entered row, unknown type or missing tare.
   status=f'IF($B{r}="","",IF(OR($L{r}="",$L{r}=0),"",IF($P{r}="",IF($R{r}="قدیمی","وزن قدیمی P01","نوع سبد را انتخاب کنید"),IF(COUNTIF(Sync!$AX$6:$AX$35,$P{r})<>1,"نوع سبد نامعتبر یا تکراری",IF(IFERROR(VLOOKUP($P{r},Sync!$AX$6:$AY$35,2,0),"")="","وزن تار را در INFO وارد کنید",IF(IFERROR(VLOOKUP($P{r},Sync!$AX$6:$AY$35,2,0),-1)<0,"وزن تار نامعتبر",""))))))'
   f=f'IF(OR($B{r}="",$K{r}="",AND($Q{r}<>"",$Q{r}<>"وزن قدیمی P01")),"",($K{r}-IF(OR($L{r}="",$L{r}=0),0,$L{r}*IF($P{r}="",IFERROR(VLOOKUP("P01",ParamTbl,3,0),1.5),VLOOKUP($P{r},Sync!$AX$6:$AY$35,2,0))))*(1-IF(IFERROR(VLOOKUP($E{r},OpTbl,3,0),"")="خرید",IFERROR(VLOOKUP("P02",ParamTbl,3,0),1),0)/100))'
   row=replace_cell(row,f'M{r}',cell(f'M{r}',formula=f,style=3))
   return append(row,cell(f'P{r}',style=5)+cell(f'Q{r}',formula=status,style=3)+(cell(f'R{r}',text='قدیمی') if r in old else ''))
  return row
 xml=rows(xml,process)
 xml=xml.replace('ref="A1:O2007"','ref="A1:R2007"',1)
 # A fixed named range includes empty slots, so later additions in INFO remain selectable.
 validation='<dataValidation type="list" allowBlank="1" showErrorMessage="1" errorTitle="نوع سبد نامعتبر" error="نوع سبد را از فهرست انتخاب کنید." sqref="P9:P2007"><formula1>BasketTypes</formula1></dataValidation>'
 xml,n=re.subn(r'(<dataValidations\b[^>]*count=")(\d+)("[^>]*>)',lambda m:m.group(1)+str(int(m.group(2))+1)+m.group(3)+validation,xml,count=1);assert n==1
 # Append narrow/hidden columns without moving any existing references.
 xml=xml.replace('</cols>','<col min="16" max="16" width="24" customWidth="1"/><col min="17" max="17" width="32" customWidth="1"/><col min="18" max="18" width="12" hidden="1" customWidth="1"/></cols>',1)
 return xml

def baskol_sync(xml):
 def process(r,row):
  if r==5:return append(row,cell('AX5',text='Basket type')+cell('AY5',text='Tare kg'))
  if 6<=r<=35:
   n=r+1
   return append(row,cell(f'AX{r}',formula=f'IF([1]Baskets!$B${n}="","",[1]Baskets!$B${n})')+cell(f'AY{r}',formula=f'IF(ISNUMBER([1]Baskets!$F${n}),[1]Baskets!$F${n},"")'))
  return row
 return rows(xml,process).replace('ref="A1:AV565"','ref="A1:AY565"',1)

def baskol_wb(xml):
 assert 'name="BasketTypes"' not in xml
 return recalc(xml).replace('</definedNames>','<definedName name="BasketTypes">Sync!$AX$6:$AX$35</definedName></definedNames>',1)
writebook(p,{'Baskol':baskol,'Sync':baskol_sync,'xl/workbook.xml':baskol_wb})

p=BASE/'Purchases.xlsx'
def purchase_sync(xml):
 def process(r,row):
  if r==5:return append(row,cell('AT5',text='Baskol net weight kg'))
  if 6<=r<=2005:
   n=r+3
   return append(row,cell(f'AT{r}',formula=f'IF([2]Baskol!$M${n}="","",[2]Baskol!$M${n})'))
  return row
 return rows(xml,process).replace('ref="A1:AS2005"','ref="A1:AT2005"',1)
def purchase_calc(xml):
 def process(r,row):
  if 9<=r<=2008:
   n=r-3;f=f'IF(OR($G{r}="",Sync!$AT{n}=""),"",IF($H{r}="مرجوعی خرید",-1,1)*Sync!$AT{n})'
   return replace_cell(row,f'N{r}',cell(f'N{r}',formula=f))
  return row
 return rows(xml,process)
writebook(p,{'Sync':purchase_sync,'Calc':purchase_calc,'xl/workbook.xml':recalc})
print('Updated Baskol.xlsx and Purchases.xlsx; legacy rows:',len(old))
