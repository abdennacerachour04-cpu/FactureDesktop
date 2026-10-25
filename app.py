# -*- coding: utf-8 -*-
"""Facture Desktop - local/offline starter application. Python 3.10+; SQLite; no Excel runtime."""
import os, sys, sqlite3, datetime, tkinter as tk, json, re
from tkinter import ttk, messagebox, filedialog
from pathlib import Path

APP_NAME = 'Facture Desktop'
BASE = Path(os.environ.get('LOCALAPPDATA', str(Path.home()))) / 'FactureDesktop'
BASE.mkdir(parents=True, exist_ok=True)
DB = BASE / 'facture_local.db'
COMPANIES_FILE = BASE / 'companies.json'

SCHEMA = '''
PRAGMA foreign_keys=ON;
CREATE TABLE IF NOT EXISTS companies(id INTEGER PRIMARY KEY, name TEXT NOT NULL, activity TEXT, address TEXT, rc TEXT, nif TEXT, ai TEXT, nis TEXT, phone TEXT, email TEXT);
CREATE TABLE IF NOT EXISTS customers(id INTEGER PRIMARY KEY, code TEXT UNIQUE, name TEXT NOT NULL, activity TEXT, address TEXT, rc TEXT, nif TEXT, ai TEXT, nis TEXT, phone TEXT, email TEXT, opening_balance REAL DEFAULT 0);
CREATE TABLE IF NOT EXISTS suppliers(id INTEGER PRIMARY KEY, code TEXT UNIQUE, name TEXT NOT NULL, activity TEXT, address TEXT, rc TEXT, nif TEXT, ai TEXT, nis TEXT, phone TEXT, email TEXT, opening_balance REAL DEFAULT 0);
CREATE TABLE IF NOT EXISTS warehouses(id INTEGER PRIMARY KEY, name TEXT UNIQUE NOT NULL, address TEXT);
CREATE TABLE IF NOT EXISTS products(id INTEGER PRIMARY KEY, code TEXT UNIQUE NOT NULL, name TEXT NOT NULL, unit TEXT, purchase_price REAL DEFAULT 0, sale_price REAL DEFAULT 0, vat REAL DEFAULT 19, margin REAL DEFAULT 0);
CREATE TABLE IF NOT EXISTS invoices(id INTEGER PRIMARY KEY, kind TEXT NOT NULL CHECK(kind IN ('SALE','PURCHASE')), number TEXT NOT NULL, date TEXT NOT NULL, partner_type TEXT NOT NULL, partner_id INTEGER, warehouse_id INTEGER, payment_mode TEXT DEFAULT 'ESPECE', discount REAL DEFAULT 0, stamp REAL DEFAULT 0, vat_recoverable INTEGER DEFAULT 1, notes TEXT, UNIQUE(kind,number));
CREATE TABLE IF NOT EXISTS invoice_lines(id INTEGER PRIMARY KEY, invoice_id INTEGER NOT NULL REFERENCES invoices(id) ON DELETE CASCADE, product_id INTEGER, code TEXT, designation TEXT NOT NULL, unit TEXT, qty REAL NOT NULL DEFAULT 1, unit_price REAL NOT NULL DEFAULT 0, vat REAL NOT NULL DEFAULT 19, amount_ht REAL NOT NULL DEFAULT 0, amount_vat REAL NOT NULL DEFAULT 0);
CREATE TABLE IF NOT EXISTS payments(id INTEGER PRIMARY KEY, date TEXT NOT NULL, partner_type TEXT NOT NULL, partner_id INTEGER, direction TEXT NOT NULL CHECK(direction IN ('IN','OUT')), amount REAL NOT NULL, mode TEXT, reference TEXT, notes TEXT);
CREATE TABLE IF NOT EXISTS stock_moves(id INTEGER PRIMARY KEY, date TEXT NOT NULL, product_id INTEGER NOT NULL REFERENCES products(id), warehouse_id INTEGER, invoice_id INTEGER, qty_in REAL DEFAULT 0, qty_out REAL DEFAULT 0, note TEXT);
'''

def connect(db_path=DB):
    c=sqlite3.connect(db_path); c.row_factory=sqlite3.Row; c.execute('PRAGMA foreign_keys=ON'); c.executescript(SCHEMA); return c

def money(v):
    try: return f'{float(v or 0):,.2f}'.replace(',', ' ')
    except: return '0.00'

def seed_from_excel(path, db_path=DB):
    """Optional one-time import of master data and historical sales/purchases from XLSM."""
    try:
        import openpyxl
        wb=openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception as e:
        raise RuntimeError('تعذر قراءة الملف. تأكد أنه ملف XLSM سليم وأن openpyxl مثبت.') from e
    c=connect(db_path); cur=c.cursor(); imported=[]
    def rows(sheet):
        if sheet not in wb.sheetnames: return []
        return list(wb[sheet].iter_rows(values_only=True))
    # Master data only; no fabricated records. Existing codes are ignored.
    for r in rows('CLIENT')[1:]:
        if len(r)>1 and r[0] and r[1]:
            cur.execute('INSERT OR IGNORE INTO customers(code,name,activity,address,rc,nif,ai,nis) VALUES(?,?,?,?,?,?,?,?)', tuple(str(x) if x is not None else None for x in (r[0],r[1],*(list(r[2:8])))))
    if cur.rowcount: imported.append('العملاء')
    for r in rows('FOURNISSEUR')[1:]:
        if len(r)>1 and r[0] and r[1]:
            cur.execute('INSERT OR IGNORE INTO suppliers(code,name,activity,address,rc,nif,ai,nis) VALUES(?,?,?,?,?,?,?,?)', tuple(str(x) if x is not None else None for x in (r[0],r[1],*(list(r[2:8])))))
    if cur.rowcount: imported.append('الموردون')
    for r in rows('DESIGNATIONS')[1:]:
        if len(r)>1 and r[0] and r[1]:
            try: cur.execute('INSERT OR IGNORE INTO products(code,name,purchase_price,unit,vat,margin,sale_price) VALUES(?,?,?,?,?,?,?)',(str(r[0]),str(r[1]),float(r[2] or 0),str(r[3] or ''),float(r[4] or 0.19)*100,float(r[5] or 0),float(r[6] or 0)))
            except (ValueError,TypeError): pass
    if cur.rowcount: imported.append('المنتجات')
    for r in rows('DEPOT')[1:]:
        if r and r[0]: cur.execute('INSERT OR IGNORE INTO warehouses(name,address) VALUES(?,?)',(str(r[0]),str(r[4] or '')))
    if cur.rowcount: imported.append('المستودعات')
    c.commit(); c.close(); wb.close()
    return imported

class App(tk.Tk):
    def __init__(self, company_name='الشركة', db_path=DB, on_switch=None):
        super().__init__(); self.company_name=company_name; self.db_path=Path(db_path); self.on_switch=on_switch; self.title(APP_NAME+' - '+company_name); self.geometry('1200x760'); self.minsize(980,620); self.configure(bg='#f2f5f9')
        self.style=ttk.Style(self); self.style.theme_use('clam'); self.style.configure('Treeview',rowheight=29,font=('Segoe UI',10)); self.style.configure('Treeview.Heading',font=('Segoe UI',10,'bold')); self.style.configure('TButton',padding=7); self.style.configure('TNotebook.Tab',padding=(16,9),font=('Segoe UI',10,'bold'))
        self.db=connect(self.db_path); self._top(); self.nb=ttk.Notebook(self); self.nb.pack(fill='both',expand=True,padx=14,pady=(0,14))
        self.make_dashboard(); self.make_master_tab('العملاء','customers',[('code','الرمز'),('name','الاسم'),('activity','النشاط'),('address','العنوان'),('phone','الهاتف'),('email','البريد')])
        self.make_master_tab('الموردون','suppliers',[('code','الرمز'),('name','الاسم'),('activity','النشاط'),('address','العنوان'),('phone','الهاتف'),('email','البريد')])
        self.make_master_tab('المنتجات','products',[('code','المرجع'),('name','التسمية'),('unit','الوحدة'),('purchase_price','سعر الشراء'),('sale_price','سعر البيع'),('vat','TVA %')])
        self.make_master_tab('المستودعات','warehouses',[('name','المستودع'),('address','العنوان')])
        self.make_invoice_tab(); self.make_payments_tab(); self.make_stock_tab()
        self.refresh_all()
    def _top(self):
        f=tk.Frame(self,bg='#14243a',height=112); f.pack(fill='x'); f.pack_propagate(False)
        head=tk.Frame(f,bg='#14243a'); head.pack(fill='x',pady=(7,3))
        tk.Label(head,text='FACTURE DESKTOP',font=('Segoe UI',18,'bold'),fg='white',bg='#14243a').pack(side='left',padx=18)
        tk.Label(head,text='إدارة الفواتير والمخزون محليًا • دون الحاجة إلى Excel',font=('Segoe UI',11),fg='#c8d7e8',bg='#14243a').pack(side='right',padx=18)
        actions=tk.Frame(f,bg='#14243a'); actions.pack(fill='x',padx=12)
        tk.Label(actions,text='الشركة: '+self.company_name,fg='white',bg='#14243a',font=('Segoe UI',10,'bold')).pack(side='right',padx=8)
        ttk.Button(actions,text='تغيير الشركة',command=self.switch_company).pack(side='right',padx=3)
        ttk.Button(actions,text='استيراد FACTURE XLSM',command=self.import_excel).pack(side='right',padx=3)
        ttk.Button(actions,text='تصدير Excel',command=self.export_excel).pack(side='right',padx=3)
        ttk.Button(actions,text='تصدير PDF',command=self.export_pdf).pack(side='right',padx=3)
        ttk.Button(actions,text='نسخة احتياطية',command=self.backup_db).pack(side='right',padx=3)
    def _export_tables(self):
        return {
            'العملاء': ('SELECT code AS الرمز,name AS الاسم,activity AS النشاط,address AS العنوان,phone AS الهاتف,email AS البريد,opening_balance AS رصيد_البداية FROM customers ORDER BY code',),
            'الموردون': ('SELECT code AS الرمز,name AS الاسم,activity AS النشاط,address AS العنوان,phone AS الهاتف,email AS البريد,opening_balance AS رصيد_البداية FROM suppliers ORDER BY code',),
            'المنتجات': ('SELECT code AS المرجع,name AS التسمية,unit AS الوحدة,purchase_price AS سعر_الشراء,sale_price AS سعر_البيع,vat AS TVA,margin AS الهامش FROM products ORDER BY code',),
            'المستودعات': ('SELECT name AS المستودع,address AS العنوان FROM warehouses ORDER BY name',),
            'الفواتير': ("SELECT i.kind AS النوع,i.number AS رقم_الفاتورة,i.date AS التاريخ,i.partner_type AS نوع_الطرف,i.partner_id AS معرف_الطرف,i.discount AS الخصم,i.stamp AS الطابع,COALESCE((SELECT SUM(amount_ht+amount_vat) FROM invoice_lines l WHERE l.invoice_id=i.id),0)-i.discount+i.stamp AS الإجمالي FROM invoices i ORDER BY i.date,i.number",),
            'تفاصيل الفواتير': ('SELECT invoice_id AS معرف_الفاتورة,code AS المرجع,designation AS التسمية,unit AS الوحدة,qty AS الكمية,unit_price AS سعر_الوحدة,vat AS TVA,amount_ht AS HT,amount_vat AS مبلغ_TVA FROM invoice_lines ORDER BY invoice_id,id',),
            'التسديدات': ('SELECT date AS التاريخ,partner_type AS نوع_الطرف,partner_id AS معرف_الطرف,direction AS الحركة,amount AS المبلغ,mode AS طريقة_الدفع,reference AS المرجع,notes AS ملاحظات FROM payments ORDER BY date,id',),
            'حركات المخزون': ("SELECT m.date AS التاريخ,p.code AS المرجع,p.name AS المنتج,w.name AS المستودع,m.qty_in AS دخول,m.qty_out AS خروج,m.note AS ملاحظات FROM stock_moves m JOIN products p ON p.id=m.product_id LEFT JOIN warehouses w ON w.id=m.warehouse_id ORDER BY m.date,m.id",),
        }
    def export_excel(self):
        path=filedialog.asksaveasfilename(title='حفظ ملف Excel',defaultextension='.xlsx',filetypes=[('Excel Workbook','*.xlsx')],initialfile='FactureDesktop_Export.xlsx')
        if not path: return
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font, PatternFill, Alignment
            from openpyxl.utils import get_column_letter
            wb=Workbook(); wb.remove(wb.active)
            for sheet_name,(query,) in self._export_tables().items():
                ws=wb.create_sheet(sheet_name[:31]); cur=self.db.execute(query); headers=[d[0] for d in cur.description]
                ws.append(headers)
                for row in cur.fetchall(): ws.append([row[h] for h in headers])
                ws.freeze_panes='A2'; ws.auto_filter.ref=ws.dimensions
                for cell in ws[1]: cell.font=Font(bold=True,color='FFFFFF'); cell.fill=PatternFill('solid',fgColor='14243A'); cell.alignment=Alignment(horizontal='center')
                for col in ws.columns:
                    letter=get_column_letter(col[0].column); width=min(max(max(len(str(c.value or '')) for c in col)+2,12),36); ws.column_dimensions[letter].width=width
            wb.save(path); messagebox.showinfo('تم التصدير','تم تصدير بيانات البرنامج إلى Excel بنجاح.');
        except Exception as e: messagebox.showerror('خطأ في التصدير',str(e))
    def export_pdf(self):
        path=filedialog.asksaveasfilename(title='حفظ تقرير PDF',defaultextension='.pdf',filetypes=[('PDF','*.pdf')],initialfile='FactureDesktop_Report.pdf')
        if not path: return
        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import A4, landscape
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.enums import TA_RIGHT, TA_CENTER
            from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
            from reportlab.pdfbase import pdfmetrics
            from reportlab.pdfbase.ttfonts import TTFont
            import os
            font_paths=[r'C:\Windows\Fonts\arial.ttf',r'C:\Windows\Fonts\tahoma.ttf','/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf']
            font_path=next((x for x in font_paths if os.path.exists(x)),None)
            if font_path: pdfmetrics.registerFont(TTFont('AppArabic',font_path)); font='AppArabic'
            else: font='Helvetica'
            try:
                import arabic_reshaper
                from bidi.algorithm import get_display
                def rtl(v):
                    v=str(v if v is not None else '')
                    try: return get_display(arabic_reshaper.reshape(v))
                    except Exception: return v
            except ImportError: rtl=lambda v:str(v if v is not None else '')
            doc=SimpleDocTemplate(path,pagesize=landscape(A4),rightMargin=22,leftMargin=22,topMargin=28,bottomMargin=25)
            styles=getSampleStyleSheet(); title_style=ParagraphStyle('ArabicTitle',parent=styles['Title'],fontName=font,fontSize=16,alignment=TA_CENTER,leading=22); normal=ParagraphStyle('Arabic',parent=styles['BodyText'],fontName=font,fontSize=7,alignment=TA_RIGHT,leading=9)
            story=[Paragraph(rtl('FACTURE DESKTOP - تقرير البيانات'),title_style),Spacer(1,10)]
            for ix,(sheet_name,(query,)) in enumerate(self._export_tables().items()):
                cur=self.db.execute(query); headers=[d[0] for d in cur.description]; rows=cur.fetchall()
                story.append(Paragraph(rtl(sheet_name),ParagraphStyle('section'+str(ix),parent=styles['Heading2'],fontName=font,alignment=TA_RIGHT,fontSize=11)))
                data=[[Paragraph(rtl(h),normal) for h in headers]]
                for row in rows:
                    data.append([Paragraph(rtl(str(row[h]) if row[h] is not None else ''),normal) for h in headers])
                if len(data)==1: data.append([Paragraph(rtl('لا توجد بيانات'),normal)]+[Paragraph('',normal) for _ in headers[1:]])
                table=Table(data,repeatRows=1,hAlign='RIGHT')
                table.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#14243a')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('GRID',(0,0),(-1,-1),0.35,colors.HexColor('#b8c4d3')),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('LEFTPADDING',(0,0),(-1,-1),4),('RIGHTPADDING',(0,0),(-1,-1),4),('TOPPADDING',(0,0),(-1,-1),4),('BOTTOMPADDING',(0,0),(-1,-1),4)]))
                story.extend([table,Spacer(1,12)])
            doc.build(story); messagebox.showinfo('تم التصدير','تم إنشاء تقرير PDF. إذا ظهر النص العربي بشكل غير صحيح، تأكد من وجود خط عربي وتبعيات PDF في نسخة البناء.')
        except Exception as e: messagebox.showerror('خطأ في PDF',str(e))
    def import_excel(self):
        path=filedialog.askopenfilename(title='اختر ملف Excel الأصلي',filetypes=[('Excel Macro-Enabled Workbook','*.xlsm'),('Excel Workbook','*.xlsx')])
        if not path: return
        if not messagebox.askyesno('تأكيد الاستيراد','سيتم استيراد البيانات الأساسية فقط حاليًا (العملاء والموردون والمنتجات والمستودعات)، مع تجاهل السجلات ذات الرموز الموجودة. لا تُنقل وحدات VBA أو استعلامات Power Query إلى التطبيق تلقائيًا. هل تريد المتابعة؟'): return
        try:
            imported=seed_from_excel(path, self.db_path)
            self.refresh_all()
            messagebox.showinfo('انتهى الاستيراد','تم فحص الملف واستيراد البيانات الأساسية.\nالجداول التي أضيفت إليها بيانات: '+(', '.join(imported) if imported else 'لم تُضف سجلات جديدة (قد تكون فارغة أو مستوردة سابقًا).'))
        except Exception as e: messagebox.showerror('فشل الاستيراد',str(e))
    def switch_company(self):
        if messagebox.askyesno('تغيير الشركة','سيتم إغلاق ملف الشركة الحالي والعودة إلى قائمة الشركات. تأكد من حفظ العمل.'): 
            try: self.db.commit(); self.db.close()
            except Exception: pass
            self.destroy()
            if self.on_switch: self.on_switch()

    def backup_db(self):
        path=filedialog.asksaveasfilename(title='حفظ نسخة احتياطية',defaultextension='.db',filetypes=[('SQLite database','*.db')],initialfile='facture_local_backup.db')
        if not path: return
        try:
            import shutil
            self.db.commit(); shutil.copy2(self.db_path,path); messagebox.showinfo('نسخة احتياطية','تم حفظ نسخة قاعدة البيانات بنجاح.')
        except Exception as e: messagebox.showerror('خطأ',str(e))

    def tab(self,title):
        f=ttk.Frame(self.nb,padding=12); self.nb.add(f,text=title); return f
    def make_dashboard(self):
        f=self.tab('الرئيسية'); tk.Label(f,text='لوحة التحكم',font=('Segoe UI',20,'bold'),bg='#f2f5f9',fg='#14243a').pack(anchor='e',pady=(5,15))
        self.cards=tk.Frame(f,bg='#f2f5f9'); self.cards.pack(fill='x')
        self.kpi={}
        for key,title in [('customers','العملاء'),('suppliers','الموردون'),('products','المنتجات'),('invoices','الفواتير')]:
            card=tk.Frame(self.cards,bg='white',highlightbackground='#dce4ee',highlightthickness=1,padx=22,pady=18); card.pack(side='right',fill='x',expand=True,padx=6)
            tk.Label(card,text=title,font=('Segoe UI',11),bg='white',fg='#65758b').pack(anchor='e')
            v=tk.Label(card,text='0',font=('Segoe UI',25,'bold'),bg='white',fg='#1664a5'); v.pack(anchor='e',pady=(8,0)); self.kpi[key]=v
        tk.Label(f,text='تنبيه: هذه نسخة أولية. راجع أرصدة البداية ونسب TVA قبل اعتمادها محاسبيًا.',font=('Segoe UI',11),bg='#f2f5f9',fg='#9a5b00').pack(anchor='e',pady=25)
    def make_master_tab(self,title,table,fields):
        f=self.tab(title); bar=tk.Frame(f); bar.pack(fill='x',pady=(0,10)); tk.Label(bar,text=title,font=('Segoe UI',17,'bold')).pack(side='right')
        ttk.Button(bar,text='＋ إضافة',command=lambda:self.edit_record(table,fields)).pack(side='left',padx=3)
        ttk.Button(bar,text='تعديل المحدد',command=lambda:self.edit_selected(tree,table,fields)).pack(side='left',padx=3)
        ttk.Button(bar,text='حذف المحدد',command=lambda:self.delete_selected(tree,table)).pack(side='left',padx=3)
        ttk.Button(bar,text='تحديث',command=lambda:self.fill_tree(tree,table,[x[0] for x in fields])).pack(side='left',padx=3)
        cols=[x[0] for x in fields]; tree=ttk.Treeview(f,columns=cols,show='headings');
        for col,label in fields: tree.heading(col,text=label); tree.column(col,width=150,anchor='center')
        tree.pack(fill='both',expand=True); setattr(self,'tree_'+table,tree)
    def edit_record(self,table,fields,record=None):
        win=tk.Toplevel(self); win.title('إضافة / تعديل'); win.geometry('460x600'); win.transient(self); win.grab_set()
        entries={};
        for col,label in fields:
            row=tk.Frame(win,padx=12,pady=5); row.pack(fill='x'); tk.Label(row,text=label,width=15,anchor='e').pack(side='right'); ent=ttk.Entry(row); ent.pack(side='right',fill='x',expand=True,padx=8); ent.insert(0,'' if record is None or record[col] is None else str(record[col])); entries[col]=ent
        def save():
            vals={k:e.get().strip() for k,e in entries.items()}
            if 'name' in vals and not vals['name']: messagebox.showerror('تنبيه','الاسم مطلوب'); return
            if 'code' in vals and not vals['code']: messagebox.showerror('تنبيه','الرمز مطلوب'); return
            try:
                if record is None:
                    q=f"INSERT INTO {table} ({','.join(vals)}) VALUES ({','.join('?' for _ in vals)})"; self.db.execute(q,list(vals.values()))
                else:
                    self.db.execute(f"UPDATE {table} SET {','.join(k+'=?' for k in vals)} WHERE id=?",list(vals.values())+[record['id']])
                self.db.commit(); win.destroy(); self.refresh_all()
            except sqlite3.IntegrityError as e: messagebox.showerror('خطأ','الرمز موجود مسبقًا أو البيانات غير صالحة.\n'+str(e))
        ttk.Button(win,text='حفظ',command=save).pack(pady=12)
    def edit_selected(self,tree,table,fields):
        s=tree.selection()
        if not s: messagebox.showinfo('تنبيه','حدد سجلًا أولًا'); return
        rid=tree.item(s[0])['tags'][0]; r=self.db.execute(f'SELECT * FROM {table} WHERE id=?',(rid,)).fetchone()
        self.edit_record(table,fields,r)
    def delete_selected(self,tree,table):
        s=tree.selection()
        if not s: messagebox.showinfo('تنبيه','حدد سجلًا أولًا'); return
        if messagebox.askyesno('تأكيد','هل تريد حذف السجل المحدد؟'):
            try: self.db.execute(f"DELETE FROM {table} WHERE id=?",(tree.item(s[0])['tags'][0],)); self.db.commit(); self.refresh_all()
            except sqlite3.IntegrityError: messagebox.showerror('تعذر الحذف','السجل مرتبط بفواتير أو حركات مخزون.')
    def fill_tree(self,tree,table,cols):
        tree.delete(*tree.get_children())
        for r in self.db.execute(f'SELECT id,{",".join(cols)} FROM {table} ORDER BY id DESC'):
            tree.insert('', 'end', values=[r[c] if r[c] is not None else '' for c in cols],tags=(str(r['id']),))
    def make_invoice_tab(self):
        f=self.tab('الفواتير'); bar=tk.Frame(f); bar.pack(fill='x'); tk.Label(bar,text='الفواتير',font=('Segoe UI',17,'bold')).pack(side='right')
        ttk.Button(bar,text='＋ فاتورة بيع',command=lambda:self.new_invoice('SALE')).pack(side='left',padx=4); ttk.Button(bar,text='＋ فاتورة شراء',command=lambda:self.new_invoice('PURCHASE')).pack(side='left',padx=4); ttk.Button(bar,text='حذف المحدد',command=lambda:self.delete_selected(tree,'invoices')).pack(side='left',padx=4)
        cols=('kind','number','date','partner_type','partner_id','total'); tree=ttk.Treeview(f,columns=cols,show='headings')
        for c,t in zip(cols,['النوع','رقم الفاتورة','التاريخ','الطرف','معرّف الطرف','الإجمالي']): tree.heading(c,text=t); tree.column(c,width=130,anchor='center')
        tree.pack(fill='both',expand=True,pady=10); self.tree_invoices=tree
    def new_invoice(self,kind):
        win=tk.Toplevel(self); win.title('فاتورة بيع' if kind=='SALE' else 'فاتورة شراء'); win.geometry('760x650'); win.transient(self); win.grab_set()
        fields=[('number','رقم الفاتورة'),('date','التاريخ YYYY-MM-DD'),('partner_id','رقم العميل/المورد'),('warehouse_id','رقم المستودع'),('product_id','رقم المنتج'),('qty','الكمية'),('unit_price','سعر الوحدة HT'),('vat','TVA %'),('discount','الخصم'),('stamp','الطابع')]; ents={}
        defaults={'number':str(self.db.execute("SELECT COUNT(*) FROM invoices WHERE kind=?",(kind,)).fetchone()[0]+1),'date':datetime.date.today().isoformat(),'qty':'1','vat':'19','discount':'0','stamp':'0'}
        for k,label in fields:
            row=tk.Frame(win,padx=12,pady=5); row.pack(fill='x'); tk.Label(row,text=label,width=22,anchor='e').pack(side='right'); e=ttk.Entry(row); e.pack(side='right',fill='x',expand=True,padx=8); e.insert(0,defaults.get(k,'')); ents[k]=e
        total=tk.Label(win,text='الإجمالي: 0.00 دج',font=('Segoe UI',14,'bold')); total.pack(pady=8)
        def calc(*_):
            try:
                ht=float(ents['qty'].get() or 0)*float(ents['unit_price'].get() or 0); vat=ht*float(ents['vat'].get() or 0)/100; total.config(text='الإجمالي: '+money(ht+vat-float(ents['discount'].get() or 0)+float(ents['stamp'].get() or 0))+' دج')
            except: total.config(text='الإجمالي: —')
        for k in ('qty','unit_price','vat','discount','stamp'): ents[k].bind('<KeyRelease>',calc)
        def save():
            try:
                vals={k:e.get().strip() for k,e in ents.items()}; qty=float(vals['qty']); price=float(vals['unit_price']); vatp=float(vals['vat']); disc=float(vals['discount']); stamp=float(vals['stamp']); ht=qty*price; tv=ht*vatp/100
                partner_type='customer' if kind=='SALE' else 'supplier'
                cur=self.db.execute('INSERT INTO invoices(kind,number,date,partner_type,partner_id,warehouse_id,discount,stamp) VALUES(?,?,?,?,?,?,?,?)',(kind,vals['number'],vals['date'],partner_type,int(vals['partner_id']) if vals['partner_id'] else None,int(vals['warehouse_id']) if vals['warehouse_id'] else None,disc,stamp)); iid=cur.lastrowid
                pr=self.db.execute('SELECT code,name,unit FROM products WHERE id=?',(int(vals['product_id']),)).fetchone() if vals['product_id'] else None
                self.db.execute('INSERT INTO invoice_lines(invoice_id,product_id,code,designation,unit,qty,unit_price,vat,amount_ht,amount_vat) VALUES(?,?,?,?,?,?,?,?,?,?)',(iid,int(vals['product_id']) if vals['product_id'] else None,pr['code'] if pr else '',pr['name'] if pr else 'بند يدوي',pr['unit'] if pr else '',qty,price,vatp,ht,tv))
                if vals['product_id']:
                    self.db.execute('INSERT INTO stock_moves(date,product_id,warehouse_id,invoice_id,qty_in,qty_out,note) VALUES(?,?,?,?,?,?,?)',(vals['date'],int(vals['product_id']),int(vals['warehouse_id']) if vals['warehouse_id'] else None,iid,qty if kind=='PURCHASE' else 0,qty if kind=='SALE' else 0,'فاتورة '+vals['number']))
                self.db.commit(); win.destroy(); self.refresh_all()
            except Exception as e: messagebox.showerror('لم تحفظ الفاتورة','تحقق من الأرقام والتاريخ ومعرّفات العميل/المورد والمستودع والمنتج.\n'+str(e))
        ttk.Button(win,text='حفظ الفاتورة',command=save).pack(pady=12)
    def make_payments_tab(self):
        f=self.tab('التسديدات'); bar=tk.Frame(f); bar.pack(fill='x'); tk.Label(bar,text='التسديدات',font=('Segoe UI',17,'bold')).pack(side='right'); ttk.Button(bar,text='＋ تسديد جديد',command=self.new_payment).pack(side='left')
        cols=('date','partner_type','partner_id','direction','amount','mode','reference'); tree=ttk.Treeview(f,columns=cols,show='headings')
        for c,t in zip(cols,['التاريخ','نوع الطرف','رقم الطرف','الحركة','المبلغ','الطريقة','المرجع']): tree.heading(c,text=t); tree.column(c,width=130,anchor='center')
        tree.pack(fill='both',expand=True,pady=10); self.tree_payments=tree
    def new_payment(self):
        win=tk.Toplevel(self); win.title('تسجيل تسديد'); win.geometry('440x430'); ents={}; options={'partner_type':['customer','supplier'],'direction':['IN','OUT'],'mode':['ESPECE','VERSMENT','CHEQUE','ATEREM']}
        for k,label in [('date','التاريخ'),('partner_type','نوع الطرف'),('partner_id','رقم الطرف'),('direction','الحركة IN/OUT'),('amount','المبلغ'),('mode','طريقة الدفع'),('reference','مرجع الشيك/التحويل'),('notes','ملاحظات')]:
            row=tk.Frame(win,padx=10,pady=5); row.pack(fill='x'); tk.Label(row,text=label,width=18,anchor='e').pack(side='right')
            if k in options: e=ttk.Combobox(row,values=options[k],state='readonly'); e.set(options[k][0])
            else: e=ttk.Entry(row); e.insert(0,datetime.date.today().isoformat() if k=='date' else '')
            e.pack(side='right',fill='x',expand=True,padx=5); ents[k]=e
        def save():
            try: self.db.execute('INSERT INTO payments(date,partner_type,partner_id,direction,amount,mode,reference,notes) VALUES(?,?,?,?,?,?,?,?)',(ents['date'].get(),ents['partner_type'].get(),int(ents['partner_id'].get() or 0) or None,ents['direction'].get(),float(ents['amount'].get()),ents['mode'].get(),ents['reference'].get(),ents['notes'].get())); self.db.commit(); win.destroy(); self.refresh_all()
            except Exception as e: messagebox.showerror('خطأ','تحقق من التاريخ والمبلغ. '+str(e))
        ttk.Button(win,text='حفظ',command=save).pack(pady=12)
    def make_stock_tab(self):
        f=self.tab('المخزون'); tk.Label(f,text='أرصدة المخزون حسب المنتج والمستودع',font=('Segoe UI',17,'bold')).pack(anchor='e')
        cols=('code','name','warehouse','incoming','outgoing','balance'); tree=ttk.Treeview(f,columns=cols,show='headings')
        for c,t in zip(cols,['المرجع','المنتج','المستودع','الدخول','الخروج','الرصيد']): tree.heading(c,text=t); tree.column(c,width=150,anchor='center')
        tree.pack(fill='both',expand=True,pady=12); self.tree_stock=tree
    def refresh_all(self):
        for table in ['customers','suppliers','products','warehouses']:
            tree=getattr(self,'tree_'+table,None)
            if tree:
                fields={'customers':['code','name','activity','address','phone','email'],'suppliers':['code','name','activity','address','phone','email'],'products':['code','name','unit','purchase_price','sale_price','vat'],'warehouses':['name','address']}[table]; self.fill_tree(tree,table,fields)
        if hasattr(self,'tree_invoices'):
            t=self.tree_invoices; t.delete(*t.get_children())
            for r in self.db.execute("SELECT i.*, COALESCE((SELECT SUM(amount_ht+amount_vat) FROM invoice_lines l WHERE l.invoice_id=i.id),0)-i.discount+i.stamp total FROM invoices i ORDER BY i.id DESC"):
                t.insert('','end',values=(('بيع' if r['kind']=='SALE' else 'شراء'),r['number'],r['date'],r['partner_type'],r['partner_id'] or '',money(r['total'])),tags=(str(r['id']),))
        if hasattr(self,'tree_payments'):
            t=self.tree_payments; t.delete(*t.get_children())
            for r in self.db.execute('SELECT * FROM payments ORDER BY id DESC'): t.insert('','end',values=(r['date'],r['partner_type'],r['partner_id'] or '',r['direction'],money(r['amount']),r['mode'],r['reference'] or ''),tags=(str(r['id']),))
        if hasattr(self,'tree_stock'):
            t=self.tree_stock; t.delete(*t.get_children())
            for r in self.db.execute('''SELECT p.code,p.name,COALESCE(w.name,'—') warehouse, SUM(m.qty_in) incoming,SUM(m.qty_out) outgoing,SUM(m.qty_in)-SUM(m.qty_out) balance FROM stock_moves m JOIN products p ON p.id=m.product_id LEFT JOIN warehouses w ON w.id=m.warehouse_id GROUP BY p.id,w.id ORDER BY p.code'''):
                t.insert('','end',values=(r['code'],r['name'],r['warehouse'],r['incoming'] or 0,r['outgoing'] or 0,r['balance'] or 0))
        if hasattr(self,'kpi'):
            for k,table in [('customers','customers'),('suppliers','suppliers'),('products','products'),('invoices','invoices')]: self.kpi[k].config(text=str(self.db.execute(f'SELECT COUNT(*) FROM {table}').fetchone()[0]))

class CompanyManager(tk.Tk):
    """Company selector. Every company has a completely separate local SQLite file."""
    def __init__(self):
        super().__init__(); self.title('Facture Desktop - إدارة الشركات'); self.geometry('760x520'); self.minsize(620,420); self.configure(bg='#f2f5f9')
        tk.Label(self,text='إدارة الشركات',font=('Segoe UI',22,'bold'),bg='#f2f5f9',fg='#14243a').pack(anchor='e',padx=24,pady=(24,4))
        tk.Label(self,text='أنشئ ملفًا مستقلًا لكل شركة للحفاظ على فصل العملاء والفواتير والمخزون والتسديدات.',font=('Segoe UI',11),bg='#f2f5f9',fg='#52657b').pack(anchor='e',padx=24,pady=(0,18))
        body=tk.Frame(self,bg='white',highlightbackground='#dce4ee',highlightthickness=1); body.pack(fill='both',expand=True,padx=22,pady=(0,14))
        self.tree=ttk.Treeview(body,columns=('name','db'),show='headings',height=10); self.tree.heading('name',text='اسم الشركة'); self.tree.heading('db',text='ملف البيانات المحلي'); self.tree.column('name',width=220,anchor='e'); self.tree.column('db',width=400,anchor='w'); self.tree.pack(fill='both',expand=True,padx=12,pady=12)
        bar=tk.Frame(self,bg='#f2f5f9'); bar.pack(fill='x',padx=22,pady=(0,18))
        ttk.Button(bar,text='＋ شركة جديدة',command=self.new_company).pack(side='right',padx=4)
        ttk.Button(bar,text='فتح الشركة المحددة',command=self.open_selected).pack(side='right',padx=4)
        ttk.Button(bar,text='تحديث القائمة',command=self.refresh).pack(side='right',padx=4)
        ttk.Button(bar,text='فتح مجلد البيانات',command=self.open_data_folder).pack(side='left',padx=4)
        self.refresh()
        self.tree.bind('<Double-1>',lambda e:self.open_selected())
    def catalog(self):
        try:
            data=json.loads(COMPANIES_FILE.read_text(encoding='utf-8'))
            return data if isinstance(data,list) else []
        except Exception: return []
    def save_catalog(self,data): COMPANIES_FILE.write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf-8')
    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        for x in self.catalog(): self.tree.insert('','end',iid=x['id'],values=(x['name'],x['db']))
    def new_company(self):
        win=tk.Toplevel(self); win.title('إنشاء شركة'); win.geometry('430x180'); win.transient(self); win.grab_set()
        tk.Label(win,text='اسم الشركة / المؤسسة',font=('Segoe UI',11,'bold')).pack(anchor='e',padx=18,pady=(18,4))
        ent=ttk.Entry(win); ent.pack(fill='x',padx=18); ent.focus_set()
        def create():
            name=ent.get().strip()
            if not name: messagebox.showerror('تنبيه','أدخل اسم الشركة.'); return
            data=self.catalog(); cid=datetime.datetime.now().strftime('%Y%m%d%H%M%S%f'); safe=re.sub(r'[^A-Za-z0-9_-]+','_',name).strip('_')[:35] or 'Company'; dbpath=BASE/('company_'+safe+'_'+cid[-6:]+'.db')
            try:
                c=connect(dbpath); c.close(); data.append({'id':cid,'name':name,'db':str(dbpath)}); self.save_catalog(data); win.destroy(); self.refresh()
                messagebox.showinfo('تم إنشاء الشركة','تم إنشاء ملف بيانات مستقل للشركة. يمكنك فتحها الآن.')
            except Exception as e: messagebox.showerror('تعذر إنشاء الشركة',str(e))
        ttk.Button(win,text='إنشاء',command=create).pack(pady=15)
    def open_selected(self):
        sel=self.tree.selection()
        if not sel: messagebox.showinfo('اختيار الشركة','حدد شركة من القائمة أولًا.'); return
        item=next((x for x in self.catalog() if x['id']==sel[0]),None)
        if not item: self.refresh(); return
        path=Path(item['db'])
        try:
            c=connect(path); c.close()
            self.withdraw()
            app=App(item['name'],path,on_switch=self.return_to_manager)
            app.protocol('WM_DELETE_WINDOW',lambda:self.close_app(app))
            app.mainloop()
        except Exception as e:
            self.deiconify(); messagebox.showerror('تعذر فتح الشركة',str(e))
    def close_app(self,app):
        try: app.db.commit(); app.db.close()
        except Exception: pass
        app.destroy(); self.destroy()
    def return_to_manager(self):
        self.deiconify(); self.refresh()
    def open_data_folder(self):
        try:
            if sys.platform.startswith('win'): os.startfile(str(BASE))
            else: messagebox.showinfo('مجلد البيانات',str(BASE))
        except Exception as e: messagebox.showerror('خطأ',str(e))

def main():
    CompanyManager().mainloop()
if __name__=='__main__': main()
