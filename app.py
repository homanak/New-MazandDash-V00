from pathlib import Path
from uuid import uuid4
from flask import Flask, render_template, request, send_file, jsonify, abort
from werkzeug.utils import secure_filename
import json

ROOT=Path(__file__).resolve().parent
UPLOADS=ROOT/'uploads'; UPLOADS.mkdir(exist_ok=True)
app=Flask(__name__)
app.config['MAX_CONTENT_LENGTH']=25*1024*1024
FILES=[('Transactions.xlsx','تراکنش‌های بانکی','۳۷۸ ردیف صورت‌حساب وارد شد'),('Cheques.xlsx','چک‌های پرداختی','۴ ردیف نمونه پاک شد'),('Employees.xlsx','حقوق و دستمزد','۴ ردیف نمونه پاک شد'),('INFO.xlsx','اطلاعات پایه','وزن تار سبد و مانده اولیه سه بانک'),('Baskol.xlsx','دفتر باسکول','انتخاب سبد و محاسبه وزن تار'),('Purchases.xlsx','خرید مرکبات','خواندن وزن خالص از باسکول'),('Baskets.xlsx','دفتر سبد','فرمول‌های تعداد، فی و گزارش اصلاح شد')]
@app.get('/')
def index():
 uploads=[]
 for meta in sorted(UPLOADS.glob('*.json'), key=lambda p:p.stat().st_mtime, reverse=True):
  try: uploads.append(json.loads(meta.read_text(encoding='utf-8')))
  except (ValueError,OSError): pass
 return render_template('index.html',files=FILES,uploads=uploads)
@app.get('/download/<name>')
def download(name):
 if name=='transactions-report':return send_file(ROOT/'گزارش-ورود-تراکنش‌ها.md',as_attachment=True,download_name='transactions-import-report.md')
 if name=='tare-guide':return send_file(ROOT/'راهنمای-اتصال-وزن-سبد.md',as_attachment=True,download_name='basket-tare-guide.md')
 if name=='design-report':return send_file(ROOT/'طراحی-تغییرات-باسکول-و-فروش.md',as_attachment=True,download_name='baskol-sales-design.md')
 if name=='baskol-report':return send_file(ROOT/'گزارش-باسکول-و-سبد.md',as_attachment=True,download_name='baskol-baskets-report.md')
 if name=='report':return send_file(ROOT/'گزارش-بررسی.md',as_attachment=True,download_name='mazandasht-report.md')
 if name not in [x[0] for x in FILES]:abort(404)
 return send_file(ROOT/'Mazandasht_Tamplate V00'/name,as_attachment=True)
@app.post('/upload')
def upload():
 f=request.files.get('file')
 if not f or not f.filename:return jsonify(error='فایلی انتخاب نشده است.'),400
 suffix=Path(f.filename).suffix.lower()
 if suffix not in {'.xlsx','.xls','.csv','.pdf'}:return jsonify(error='فقط فایل‌های Excel، CSV و PDF پذیرفته می‌شوند.'),400
 key=uuid4().hex; dest=UPLOADS/(key+suffix); f.save(dest)
 meta={'id':key,'name':Path(f.filename.replace('\\','/')).name,'size':dest.stat().st_size,'suffix':suffix}
 (UPLOADS/(key+'.json')).write_text(json.dumps(meta,ensure_ascii=False),encoding='utf-8')
 return jsonify(ok=True)
@app.get('/uploads/<key>')
def uploaded(key):
 if len(key)!=32 or any(c not in '0123456789abcdef' for c in key):abort(404)
 meta=UPLOADS/(key+'.json')
 if not meta.exists():abort(404)
 info=json.loads(meta.read_text(encoding='utf-8')); return send_file(UPLOADS/(key+info['suffix']),as_attachment=True,download_name=info['name'])
@app.errorhandler(413)
def large(e):return jsonify(error='حجم فایل باید کمتر از ۲۵ مگابایت باشد.'),413
if __name__=='__main__':app.run(host='0.0.0.0',port=8000)
