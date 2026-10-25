# تقرير فحص ملف FACTURE-1.xlsm

- عدد أوراق العمل: 26
- عدد الصيغ المخزنة داخل أوراق العمل: 409
- وجود مشروع VBA: نعم (xl/vbaProject.bin موجود)
- عدد ملفات QueryTable: 10
- عدد الاتصالات المسجلة: 15

## أوراق العمل

- AJOUTE DEPOT
- SOLDE DE FOURNISSEURS
- REGLEMNTFOURNISSEUR
- EDTIONREGLEMNTF
- LISTE DE COMPTES
- EDTIONREGLEMNTC
- SOLDE CLIENT
- REGLEMNTCLIENT
- AJOUTE DESIGNATIONS
- AJOUTE FOURNISSEUR
- FOURNISSEUR
- ETATE ACHATS
- ACHATS
- FACTOURACHA
- FACTOURVENT
- STOCK
- VENTS
- CLIENT
- DEPOT
- DESIGNATIONS
- EN STOCK
- HOUSE
-   LES TABLEAUX
- MODE DE PAIEMENT
- AJOUTE CLIENT
- ETATE VENTE

## الاتصالات والاستعلامات المسجلة

- **استعلام - ‏‏CLNTE** — الأمر: `SELECT * FROM [CLNTE]` — مصدر الاتصال: `Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=CLNTE;Extended Properties=""`
- **استعلام - ‏‏DEPOT** — الأمر: `SELECT * FROM [DEPOT]` — مصدر الاتصال: `Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=DEPOT;Extended Properties=""`
- **استعلام - ‏‏DESIGNATIONS** — الأمر: `SELECT * FROM [DESIGNATIONS]` — مصدر الاتصال: `Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=DESIGNATIONS;Extended Properties=""`
- **استعلام - ‏‏EN STOCK** — الأمر: `SELECT * FROM [EN STOCK]` — مصدر الاتصال: `Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location="EN STOCK";Extended Properties=""`
- **استعلام - ‏‏ENTREES** — الأمر: `SELECT * FROM [ENTREES]` — مصدر الاتصال: `Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=ENTREES;Extended Properties=""`
- **استعلام - ‏‏ENTREES (2)** — الأمر: `SELECT * FROM [ENTREES (2)]` — مصدر الاتصال: `Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location="ENTREES (2)";Extended Properties=""`
- **استعلام - ‏‏ETATE ACHATS** — الأمر: `SELECT * FROM [ETATE ACHATS]` — مصدر الاتصال: `Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location="ETATE ACHATS";Extended Properties=""`
- **استعلام - ‏‏ETATE VENTE** — الأمر: `SELECT * FROM [ETATE VENTE]` — مصدر الاتصال: `Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location="ETATE VENTE";Extended Properties=""`
- **استعلام - ‏‏ETATE VENTE (2)** — الأمر: `SELECT * FROM [ETATE VENTE (2)]` — مصدر الاتصال: `Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location="ETATE VENTE (2)";Extended Properties=""`
- **استعلام - ‏‏FOURNISSEURS** — الأمر: `SELECT * FROM [FOURNISSEURS]` — مصدر الاتصال: `Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=FOURNISSEURS;Extended Properties=""`
- **استعلام - ‏‏REG FR** — الأمر: `SELECT * FROM [REG FR]` — مصدر الاتصال: `Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location="REG FR";Extended Properties=""`
- **استعلام - ‏‏SOLDE CLIENT** — الأمر: `SELECT * FROM [SOLDE CLIENT]` — مصدر الاتصال: `Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location="SOLDE CLIENT";Extended Properties=""`
- **استعلام - ‏‏SOMDE DE FOURNISSEURS** — الأمر: `SELECT * FROM [SOMDE DE FOURNISSEURS]` — مصدر الاتصال: `Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location="SOMDE DE FOURNISSEURS";Extended Properties=""`
- **استعلام - ‏‏SORTIES** — الأمر: `SELECT * FROM [SORTIES]` — مصدر الاتصال: `Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=SORTIES;Extended Properties=""`
- **استعلام - ‏‏SORTIES (2)** — الأمر: `SELECT * FROM [SORTIES (2)]` — مصدر الاتصال: `Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location="SORTIES (2)";Extended Properties=""`

## مهم قبل اعتماد التطبيق

- هذا التقرير يستخرج أسماء الاتصالات وأوامرها المسجلة فقط. أوامر `SELECT * FROM [اسم]` لا تكشف وحدها منطق Power Query الكامل، وقد يكون تعريف التحويل داخل بيانات Mashup المضمّنة.
- وجود `vbaProject.bin` يعني أن الملف يحتوي وحدات VBA، لكنه لا يعني أن كود VBA يعمل مباشرة داخل تطبيق Windows مستقل.
- لا يمكن تشغيل VBA أو Power Query الخاصين بـ Excel داخل البرنامج الجديد كما هما. يجب استخراج كل إجراء/استعلام وفهمه ثم إعادة برمجة نتيجته في التطبيق، واختبار الأرصدة والضرائب والفواتير على عينات مطابقة.
- الاستيراد الحالي ينقل بيانات أساسية فقط من أوراق CLIENT وFOURNISSEUR وDESIGNATIONS وDEPOT؛ لا ينقل سجل الفواتير التاريخية ولا يدّعي تنفيذ الماكرو أو الاستعلامات.
- خيارات التصدير الجديدة تصدر البيانات الموجودة داخل قاعدة البرنامج إلى XLSX أو PDF، ولا تصدر كود VBA نفسه.


## قرارات النسخة متعددة الشركات

- أضيفت شاشة مستقلة لاختيار الشركة وإنشاء شركات متعددة.
- لكل شركة ملف SQLite منفصل، لتجنب اختلاط السجلات بين المؤسسات.
- عند استيراد ملف XLSM، تُنقل البيانات الأساسية إلى قاعدة الشركة المفتوحة فقط.
- هذه الخطوة لا تعني أن جميع أكواد VBA أو تحويلات Power Query أعيدت برمجتها؛ يلزم استخراج نصوص VBA ومحتوى Mashup ومراجعة أثرها على المحاسبة قبل إعلان التطابق الكامل.
