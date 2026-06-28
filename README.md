# STC Quality Executive Dashboard — One File Version

هذه نسخة مبسطة تعمل على ملف واحد فقط: `Deviation.xlsx`.

## الملفات المطلوبة
- `app.py`
- `requirements.txt`
- `Deviation.xlsx`

## التشغيل
```bash
pip install -r requirements.txt
streamlit run app.py
```

## محتوى الداشبورد
- Total Deviations
- Unique Work Orders
- Penalty Applied
- No Penalty
- Service Affecting
- تصنيف المخالفات إلى Civil / Fiber / Safety / Other
- Top Work Orders حسب عدد المخالفات
- طبيعة المخالفات لكل Work Order
- WO Classification Heatmap
- Export Excel Summary
- Export Executive PDF Report

## ملاحظات
- التصنيف يعتمد أولًا على عمود `Designation` إذا كان موجودًا.
- إذا لم يكن `Designation` واضحًا، يتم التصنيف من خلال الكلمات المفتاحية في Category / SubCategory / DeviationName.
- لا توجد ملفات CSV إضافية مطلوبة.
