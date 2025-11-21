# ================== الاصدار الذكى (معدل) =========================
# MRP Analysis Tool Final Version with Stock Analysis and Component Order Type
# Developed by: Reda Roshdy
# Date: 17-Sep-2025
# التعديل: الاعتماد على عمود MRP Controller من ورقة المكونات مباشرةً
# ==============================================================================

# ==============================================================================
# 1. استدعاء المكتبات اللازمة
# ==============================================================================
import streamlit as st
import pandas as pd
import datetime
import re # ⬅️ إضافة هذا السطر
from io import BytesIO
import zipfile
import calendar
import plotly.express as px

# ==============================================================================
# 2. إعداد التكوين والأعمدة (تحسين جديد)
# ==============================================================================

COLUMN_NAMES = {
    "material": ["Material", "Item", "code", "Code", "المادة", "Product"],
    "material_desc": ["Material Description", "Description", "وصف"],
    "order_type": ["Order Type", "OT", "نوع الطلب", "Sales Org."],
    "component": ["Component", "Comp", "المكون"],
    "component_desc": ["Component Description", "Comp Desc", " المسمى", "وصف المكون"],
    "component_uom": ["Component UoM", "UoM", "الوحدة"],
    "component_qty": ["Component Quantity", "Qty", "كمية المكون"],
    "mrp_controller": ["MRP Contor", "MRP Controller", "مسؤول MRP"],
    "current_stock": ["Current Stock", "Stock", "المخزون الحالي", "Unrestricted"],
    "component_order_type": ["Component Order Type", "Order Category", "نوع أمر المكون", "Procurement Type"],
    "hierarchy_level": ["Hierarchy Level", "Level", "المستوى الهرمي"]
}
# ==============================================================================
# 2. تنظيف وتحويل الأرقام
# ==============================================================================
def clean_numeric(val):
    if pd.isna(val): return 0.0
    s = str(val).strip().replace(',', '')
    neg = s.endswith('-') or (s.startswith('(') and s.endswith(')'))
    s = re.sub(r'[^0-9.]', '', s)
    try: return -float(s) if neg else float(s)
    except: return 0.0

def ensure_numeric(df, colname):
    df[colname] = df.get(colname, 0.0).apply(clean_numeric)
# ==============================================================================
# 3. توحيد الوحدات (جرام → كجم)
# ==============================================================================
def normalize_units(df, qty_col):
    grams = ["g","gram","grams","gm","جرام","غ"]
    df = df.copy()
    # ⬇️ التعديل هنا: استخدام col("component_uom") بدلاً من "Component UoM"
    uom_col = col("component_uom")
    
    df[uom_col] = df[uom_col].astype(str).str.lower()
    mask = df[uom_col].isin(grams)
    
    df.loc[mask, qty_col] /= 1000
    df.loc[mask, uom_col] = "kg"
    return df
# ==============================================================================
# 3. الدوال المساعدة (Functions)
# ==============================================================================
def col(name_key):
    return COLUMN_NAMES[name_key][0]

def normalize_columns(df, column_map):
    rename_dict = {}
    for key, aliases in column_map.items():
        if isinstance(aliases, list):
            for alias in aliases:
                if alias in df.columns:
                    rename_dict[alias] = aliases[0]
        else:
            if aliases in df.columns:
                rename_dict[aliases] = aliases
    return df.rename(columns=rename_dict)

@st.cache_data
def load_and_validate_data(uploaded_file):
    try:
        xls = pd.ExcelFile(uploaded_file, engine='openpyxl')

        required_sheets = ["plan", "Component"]
        missing_sheets = [sheet for sheet in required_sheets if sheet not in xls.sheet_names]
        if missing_sheets:
            st.error(f"❌ الملف لا يحتوي على الأوراق المطلوبة: {', '.join(missing_sheets)}")
            st.stop()
            
        plan_df = normalize_columns(xls.parse("plan"), COLUMN_NAMES)
        component_df = normalize_columns(xls.parse("Component"), COLUMN_NAMES)
        # ❌ تم إلغاء قراءة ورقة MRP Contor المنفصلة
        mrp_df = pd.DataFrame() 

        # التحقق من الأعمدة الأساسية
        required_plan_columns = [col("material"), col("material_desc"), col("order_type")]
        if not all(c in plan_df.columns for c in required_plan_columns):
            st.error(f"❌ جدول الخطة لا يحتوي على الأعمدة المطلوبة: {', '.join(required_plan_columns)}")
            st.stop()

        required_component_columns = [col("material"), col("component"), col("component_qty")]
        if not all(c in component_df.columns for c in required_component_columns):
            st.error(f"❌ جدول المكونات لا يحتوي على الأعمدة المطلوبة: {', '.join(required_component_columns)}")
            st.stop()
            
        # ✅ إضافة تحقق إلزامي لعمود MRP Controller في ورقة المكونات
        if col("mrp_controller") not in component_df.columns:
            st.error(f"❌ جدول المكونات لا يحتوي على العمود المطلوب: {col('mrp_controller')}")
            st.stop()
            
# التحقق من وجود أعمدة اختيارية وإضافة الافتراضي إن لزم الأمر
        if col("current_stock") not in component_df.columns:
            component_df[col("current_stock")] = 0

        if col("component_order_type") not in component_df.columns:
            component_df[col("component_order_type")] = "غير محدد"
        
        if col("hierarchy_level") not in component_df.columns:
            component_df[col("hierarchy_level")] = "غير محدد"           
# ==============================================================================
        # 🗜️ تطبيق تنظيف البيانات وتوحيد الوحدات (جديد)
# ==============================================================================
        # 1. ضمان تحويل الكميات والمخزون إلى أرقام
        ensure_numeric(component_df, col("component_qty"))
        ensure_numeric(component_df, col("current_stock")) # الآن أصبح موجوداً بفضل السطر 119
        
        # 2. توحيد الوحدات (جرام -> كجم)
        component_df = normalize_units(component_df, col("component_qty"))
# ==============================================================================
        return plan_df, component_df, mrp_df

    except Exception as e:
        st.error(f"حدث خطأ أثناء قراءة الملف: {e}")
        st.stop()
# ==============================================================================
# 4. واجهة المستخدم الرئيسية للتطبيق
# ==============================================================================

st.set_page_config(page_title="🔥 MRP Tool", page_icon="📂", layout="wide")
st.header("📂 MRP الاصدار الذكى من برنامج تحليل واستخراج وحفظ نتائج الـ")
# دليل الاستخدام
with st.expander("📖 دليل الاستخدام"):
    st.write("""
    ### كيفية استخدام البرنامج:
    1. **حمل الملف**: اختر ملف Excel يحتوي على أوراق (plan و Component)
    2. **استخدم الفلاتر**: طبّق المرشحات لتضييق النتائج حسب احتياجك
    3. **ابحث**: استخدم خاصية البحث السريع للعثور على مكونات محددة
    4. **حلل**: راجع الرسوم البيانية والتنبيهات
    5. **صدّر**: احفظ النتائج بصيغة Excel
    """)

st.markdown("<p style='font-size:16px; font-weight:bold;'>📂 اختر ملف الخطة الشهرية Excel</p>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("", type=["xlsx"])

if uploaded_file:
    # mrp_df سيتم تمريره فارغاً
    plan_df, component_df, mrp_df = load_and_validate_data(uploaded_file)
    plan_df_orig = plan_df.copy()
    component_df_orig = component_df.copy()
    mrp_df_orig = mrp_df.copy()

    # 🚨 التصحيح 1: تعريف جداول فارغة لضمان عدم ظهور NameError في جزء التصدير
    summary_df = pd.DataFrame()
    component_bom_pivot = pd.DataFrame()

    # أي معالجة أو جداول Pivot بعد كده...

    # استخراج أعمدة التواريخ مرة واحدة
    date_cols = [c for c in plan_df.columns if isinstance(c, (datetime.datetime, pd.Timestamp))]
    
    # نسخة معالجة
    plan_df_processed = plan_df.copy()

    # 🔹 إجبار أعمدة الأكواد إنها تبقى نصوص لتفادي الفواصل

    with st.spinner("⏳ جاري معالجة البيانات وعرض النتائج..."):
        # (نفس الحسابات والجداول والرسوم البيانية الموجودة في كودك الأصلي بدون تعديل)

# ==============================================================================
        # تحويل شيت الخطة إلى شكل طويل (Plan long)
# ==============================================================================
        id_vars = [col("material"), col("material_desc"), col("order_type")]
        # تأكد أن الأعمدة موجودة في حال اختلاف الحروف الكبيرة/الصغيرة
        id_vars = [c for c in id_vars if c in plan_df.columns]
        value_vars = [c for c in plan_df.columns if c not in id_vars]

        plan_melted = plan_df.melt(
            id_vars=id_vars,
            value_vars=value_vars,
            var_name="Date",
            value_name="Planned Quantity"
        )
        plan_melted["Date"] = pd.to_datetime(plan_melted["Date"], errors='coerce')
        plan_melted = plan_melted.dropna(subset=["Planned Quantity"])
        # نتأكد من أن الكمية رقم
        plan_melted["Planned Quantity"] = pd.to_numeric(plan_melted["Planned Quantity"], errors='coerce').fillna(0)

# ==============================================================================
        # Merge المباشر (كمقياس لمقاربات سابقة) - يبقى موجود للاستعلامات الأخرى
# ==============================================================================
        merged_df = pd.merge(plan_melted, component_df, on=col("material"), how="left")
        merged_df["Required Component Quantity"] = merged_df["Planned Quantity"] * merged_df[col("component_qty")]

# ==============================================================================
        # ======= Recursive BOM Aggregation =======
# ==============================================================================
        
        # دالة تفجير تكراري مع منع الحلقات (loop prevention)
        def explode_recursive(parent_material, qty, date, comp_df, results, path):
            """
            parent_material: كود المادة الأب (Material)
            qty: الكمية الحالية (مضروبة حتى هذه النقطة)
            date: تاريخ الطلب (pd.Timestamp أو نص)
            comp_df: DataFrame مكونات
            results: قائمة لتخزين الصفوف
            path: قائمة المكونات في المسار الحالي (لتفادي الحلقات)
            """
            # العثور على الأبناء في BOM
            children = comp_df[comp_df[col("material")] == parent_material]
            if children.empty:
                return
            for _, row in children.iterrows():
                child_code = row[col("component")]
                # منع الحلقات: إذا ظهر العنصر مسبقًا في المسار، تجاهل النزول له مرة أخرى
                if child_code in path:
                    continue
                # قراءة كمية المكون (قد تكون نص)
                try:
                    per_unit = float(row.get(col("component_qty"), 0) or 0)
                except:
                    per_unit = 0.0
                child_qty = qty * per_unit
                # إضافة الصف
                results.append({
                    col("component"): child_code,
                    col("component_desc"): row.get(col("component_desc"), ""),
                    col("component_uom"): row.get(col("component_uom"), ""),
                    "Procurement Type": row.get(col("component_order_type"), ""),
                    col("mrp_controller"): row.get(col("mrp_controller"), "N/A"), # ✅ تم جلب MRP Contor من صف المكون
                    "Date": date,
                    "Required Qty": child_qty
                })
                # تكرار النزول أسفل هذا الطفل
                explode_recursive(child_code, child_qty, date, comp_df, results, path + [child_code])

        # تجهيز قائمة النتائج
        recursive_results = []

        # نفذ التفجير لكل صف في plan_melted
        for _, plan_row in plan_melted.iterrows():
            top_mat = plan_row[col("material")]
            plan_qty = plan_row["Planned Quantity"]
            order_date = plan_row["Date"]  # pd.Timestamp or NaT
            # إذا الكمية صفر نتخطى
            if plan_qty == 0:
                continue
            # نفجر من المنتج النهائي
            explode_recursive(top_mat, plan_qty, order_date, component_df, recursive_results, path=[top_mat])

        # تحويل للقيم DataFrame
        recursive_df = pd.DataFrame(recursive_results)
        if not recursive_df.empty:
            # تجميع حسب المكون والتاريخ
            agg_recursive = recursive_df.groupby(
                [col("component"), col("component_desc"), col("component_uom"), "Procurement Type", col("mrp_controller"), "Date"], # ✅ تمت إضافة MRP Contor إلى التجميع
                as_index=False
            )["Required Qty"].sum()

            # ❌ تم حذف دمج MRP Contor من mrp_df

            # تحويل التاريخ إلى نص dd mmm في العناوين لاحقاً عند pivot
            agg_recursive["Date"] = pd.to_datetime(agg_recursive["Date"], errors='coerce')

            # عمل Pivot بحيث كل تاريخ عمود
            pivot_recursive = agg_recursive.pivot_table(
                index=[col("component"), col("component_desc"), col("component_uom"), "Procurement Type", col("mrp_controller")],
                columns="Date",
                values="Required Qty",
                aggfunc="sum",
                fill_value=0
            ).reset_index()
            
            # إعادة تسمية عمود MRP Contor الناتج من التجميع
            pivot_recursive.rename(columns={col("mrp_controller"): "MRP Contor"}, inplace=True)


            # تنسيق أسماء أعمدة التاريخ لعرض dd mmm
            pivot_recursive.columns = [
                (col.strftime("%d %b") if isinstance(col, pd.Timestamp) else col) for col in pivot_recursive.columns
            ]

        else:
            pivot_recursive = pd.DataFrame(columns=[col("component"), col("component_desc"), col("component_uom"), "Procurement Type", "MRP Contor"])

# ==============================================================================
        # الملخص السريع (عرض فقط)
# ==============================================================================
        total_models = plan_df[col("material")].nunique()
        total_components = component_df[col("component")].nunique()
        total_boms = len(component_df)
        
        # ❌ حذف إحصائية المكونات بدون MRP Contor من mrp_df
        # ✅ حساب إحصائية المكونات بدون MRP Contor من ورقة المكونات مباشرة
        empty_mrp_count = component_df[col("mrp_controller")].isna().sum()


        diff_uom = component_df.groupby(col("component"))[col("component_uom")].nunique()
        diff_uom = diff_uom[diff_uom > 1]
        total_diff_uom = len(diff_uom)

        if total_diff_uom > 0:
            diff_uom_str = ", ".join(map(str, diff_uom.index))
            diff_uom_color = "red"
        else:
            diff_uom_str = "لا يوجد"
            diff_uom_color = "green"

        missing_boms = set(plan_df[col("material")]) - set(component_df[col("material")])
        total_missing_boms = len(missing_boms)
        missing_boms_html = (
            f"<span style='color:red;'>{', '.join(map(str, missing_boms))}</span>"
            if missing_boms else "<span style='color:green;'>لا يوجد</span>"
        )
    
# ==============================================================================
        # إحصائية جديدة لأنواع طلب المكونات
# ==============================================================================

        # خريطة الأكواد إلى النصوص
        order_type_map = {
            "F": "شراء",
            "E": "تصنيع"
        }

        # إضافة عمود جديد بالوصف العربي
        component_df["Order_Type_Label"] = component_df[col("component_order_type")].map(order_type_map).fillna("غير محدد")
# ==============================================================================
        # حساب الإحصائيات بعد توحيد الأعمدة
        purchase_count = component_df.loc[component_df["Order_Type_Label"] == "شراء", col("component")].nunique()        # عدد المكونات شراء
        manufacturing_count = component_df.loc[component_df["Order_Type_Label"] == "تصنيع", col("component")].nunique()  # عدد المكونات تصنيع
        undefined_count = component_df.loc[component_df["Order_Type_Label"] == "غير محدد", col("component")].nunique()   # عدد المكونات غير محددة

# -------------------------------
        # بناء جدول الملخص (Summary Sheet) للتصدير
# -------------------------------
        summary_data_rows = [
            {"المقياس": "📌 ملخص نتائج الخطة", "القيمة": "", "ملاحظات": ""},
            {"المقياس": "🟢 موديلات بوتاجاز بالخطة", "القيمة": total_models, "ملاحظات": ""},
            {"المقياس": "🔵 عدد المكونات المستخدمة", "القيمة": total_components, "ملاحظات": ""},
            {"المقياس": "🟠 إجمالي عدد مكونات الـ BOMs", "القيمة": total_boms, "ملاحظات": ""},
            
            {"المقياس": "---", "القيمة": "---", "ملاحظات": "---"},

            {"المقياس": "❌ مكونات بدون MRP Contor", "القيمة": empty_mrp_count, "ملاحظات": ""},
            {"المقياس": "⚠️ مكونات لها أكثر من وحدة", "القيمة": total_diff_uom, "ملاحظات": diff_uom_str},
            {"المقياس": "✅ منتجات بالخطة بدون BOM", "القيمة": total_missing_boms, "ملاحظات": ", ".join(map(str, missing_boms)) if missing_boms else "لا يوجد"},
            
            {"المقياس": "---", "القيمة": "---", "ملاحظات": "---"},
            
            {"المقياس": "🔹 ملخص أنواع طلب المكونات", "القيمة": "", "ملاحظات": ""},
            {"المقياس": "🛒 مكونات شراء", "القيمة": purchase_count, "ملاحظات": ""},
            {"المقياس": "🏭 مكونات تصنيع", "القيمة": manufacturing_count, "ملاحظات": ""},
            {"المقياس": "❓ مكونات غير محددة", "القيمة": undefined_count, "ملاحظات": ""},
        ]
        
        summary_df = pd.DataFrame(summary_data_rows)
        # -------------------------------
        
        st.markdown(f"""
        <div style="direction:rtl; text-align:right; font-size:20px;">
        <span style="font-size:22px; color:#1976d2;">📌 <b>ملخص نتائج الخطة </b></span>
        <br><br>
        <ul style="list-style-type:none; padding-right:0;">

          <li>🟢 <b>{total_models}</b> موديلات بوتاجاز بالخطة</li>
          <li>🔵 <b>{total_components}</b> عدد المكونات المستخدمة</li>
          <li>🟠 <b>{total_boms}</b> إجمالي عدد مكونات الـ BOMs</li>

          <li>{"❌" if empty_mrp_count>0 else "✅"} 
              <b style="color:{'red' if empty_mrp_count>0 else 'green'};">
              {empty_mrp_count}</b> مكونات بدون MRP Contor</li>

          <li>{"⚠️" if total_diff_uom>0 else "✅"} 
              <b style="color:{'red' if total_diff_uom>0 else 'green'};">
              {total_diff_uom}</b> مكونات لها أكثر من وحدة: 
              <span style="color:{diff_uom_color};">{diff_uom_str}</span>
          </li>

          <li>{"⚠️" if total_missing_boms>0 else "✅"} 
              <b style="color:{'red' if total_missing_boms>0 else 'green'};">
              {total_missing_boms}</b> منتجات موجودة بالخطة لكن بدون BOM: 
              {missing_boms_html}
          </li>

        </ul>
        </div>
        """, unsafe_allow_html=True)

        # عرض إحصائية أنواع طلب المكونات
        st.markdown(f"""
        <div style="direction:rtl; text-align:right; font-size:20px;">
        <span style="font-size:22px; color:#1976d2;">🔹 <b>ملخص أنواع طلب المكونات</b></span>
        <br><br>
        <ul style="list-style-type:none; padding-right:0;">
            <li>🛒 <b>{purchase_count}</b> مكونات شراء</li>
            <li>🏭 <b>{manufacturing_count}</b> مكونات تصنيع</li>
            <li>❓ <b>{undefined_count}</b> مكونات غير محددة</li>
        </ul>
        </div>
        """, unsafe_allow_html=True)

# ==============================================================================
        # Need_By_Date - حساب باستخدام Recursive BOM
# ==============================================================================

        # دالة تفجير تكراري مخصصة لحساب Need_By_Date (تأخذ معلومات Current Stock و Component Order Type من صف المكون)
        def explode_recursive_need(parent_material, qty, date, comp_df, results, path):
            children = comp_df[comp_df[col("material")] == parent_material]
            if children.empty:
                return
            for _, crow in children.iterrows():
                child_code = crow[col("component")]
                # منع الحلقات
                if child_code in path:
                    continue
                # قراءة الكمية لكل وحدة مع الحماية من القيم النصية
                try:
                    per_unit = float(crow.get(col("component_qty"), 0) or 0)
                except:
                    per_unit = 0.0
                child_qty = qty * per_unit

                results.append({
                    col("component"): child_code,
                    col("component_desc"): crow.get(col("component_desc"), ""),
                    col("component_uom"): crow.get(col("component_uom"), ""),
                    col("current_stock"): crow.get(col("current_stock"), 0),
                    col("component_order_type"): crow.get(col("component_order_type"), ""),
                    col("mrp_controller"): crow.get(col("mrp_controller"), "N/A"), # ✅ جلب MRP Contor من صف المكون
                    "Date": date,
                    "Required Component Quantity": child_qty
                })

                # استدعاء تكراري للطفل
                explode_recursive_need(child_code, child_qty, date, comp_df, results, path + [child_code])

        # تنفيذ التفجير لكل صف في plan_melted
        need_results = []
        for _, prow in plan_melted.iterrows():
            top_material = prow[col("material")]
            plan_qty = prow["Planned Quantity"]
            order_date = prow["Date"]
            if plan_qty == 0 or pd.isna(order_date):
                continue
            explode_recursive_need(top_material, plan_qty, order_date, component_df, need_results, path=[top_material])

        # تحويل النتائج إلى DataFrame وتجميعها
        need_df = pd.DataFrame(need_results)
        if not need_df.empty:
            # تجميع حسب المكون والتاريخ مع جمع الكميات المطلوبة الناتجة من التفجير التكراري
            result_date = need_df.groupby(
                [col("component"), col("component_desc"), col("component_uom"), col("current_stock"), col("component_order_type"), col("mrp_controller"), "Date"], # ✅ إضافة MRP Contor هنا
                as_index=False
            )["Required Component Quantity"].sum()

            # عمل Pivot بحيث كل تاريخ يصبح عمودًا
            pivot_by_date = result_date.pivot_table(
                index=[col("component"), col("component_desc"), col("component_uom"), col("current_stock"), col("component_order_type"), col("mrp_controller")], # ✅ إضافة MRP Contor هنا
                columns="Date",
                values="Required Component Quantity",
                aggfunc="sum",
                fill_value=0
            ).reset_index()

            # ❌ تم حذف دمج عمود MRP Contor من mrp_df
            # ✅ إعادة تسمية العمود
            pivot_by_date.rename(columns={col("mrp_controller"): "MRP Contor"}, inplace=True)


            # إعادة ترتيب الأعمدة
            cols = pivot_by_date.columns.tolist()
            fixed_order = [col("component"), col("component_desc"), "MRP Contor", col("component_uom"), col("current_stock"), col("component_order_type")]
            other_cols = [c for c in cols if c not in fixed_order]
            pivot_by_date = pivot_by_date[fixed_order + other_cols]

            # تنسيق أسماء الأعمدة (التواريخ تبقى dd mmm)
            pivot_by_date.columns = [
                col.strftime("%d %b") if isinstance(col, pd.Timestamp) else col
                for col in pivot_by_date.columns
            ]
        else: # 🚨 التصحيح 2: تعريف جدول فارغ في حالة عدم وجود نتائج لتفادي NameError
            pivot_by_date = pd.DataFrame(columns=[col("component"), col("component_desc"), "MRP Contor", col("component_uom"), col("current_stock"), col("component_order_type")])


# ==============================================================================
        # Need_By_Order Type - Recursive per Month + OrderType
# ==============================================================================

        def explode_recursive_order(parent_material, qty, order_type, order_date, comp_df, results, path):
            children = comp_df[comp_df[col("material")] == parent_material]
            if children.empty:
                return
            for _, crow in children.iterrows():
                child_code = crow[col("component")]
                if child_code in path:
                    continue
                try:
                    per_unit = float(crow.get(col("component_qty"), 0) or 0)
                except:
                    per_unit = 0.0
                child_qty = qty * per_unit

                results.append({
                    col("component"): child_code,
                    col("component_desc"): crow.get(col("component_desc"), ""),
                    col("component_uom"): crow.get(col("component_uom"), ""),
                    col("current_stock"): crow.get(col("current_stock"), 0),
                    col("component_order_type"): crow.get(col("component_order_type"), ""),
                    col("mrp_controller"): crow.get(col("mrp_controller"), "N/A"), # ✅ جلب MRP Contor من صف المكون
                    "Order Type": order_type,
                    "Month": pd.to_datetime(order_date).strftime("%b"),  # الشهر فقط
                    "Required Component Quantity": child_qty
                })

                explode_recursive_order(child_code, child_qty, order_type, order_date, comp_df, results, path + [child_code])

        # تنفيذ التفجير عبر الخطة كلها
        order_results = []
        for _, prow in plan_melted.iterrows():
            top_material = prow[col("material")]
            plan_qty = prow["Planned Quantity"]
            order_type = prow.get(col("order_type"), "N/A")
            order_date = prow.get("Date", None)
            if plan_qty == 0 or pd.isna(order_date):
                continue
            explode_recursive_order(top_material, plan_qty, order_type, order_date, component_df, order_results, path=[top_material])

        order_df = pd.DataFrame(order_results)

        if not order_df.empty:
            # تجميع حسب (Component + OrderType + Month)
            result_order = order_df.groupby(
                [col("component"), col("component_desc"), col("component_uom"), col("current_stock"), col("component_order_type"), col("mrp_controller"), "Order Type", "Month"], # ✅ إضافة MRP Contor هنا
                as_index=False
            )["Required Component Quantity"].sum()

            # إنشاء عمود تجميعي لكل نوع طلب وشهر
            result_order["Order_Month"] = result_order["Month"] + " (" + result_order["Order Type"] + ")"

            pivot_by_order = result_order.pivot_table(
                index=[col("component"), col("component_desc"), col("component_uom"), col("current_stock"), col("component_order_type"), col("mrp_controller")], # ✅ إضافة MRP Contor هنا
                columns="Order_Month",
                values="Required Component Quantity",
                aggfunc="sum",
                fill_value=0
            ).reset_index()

            # ❌ تم حذف دمج مع MRP Contor لو متاح
            # ✅ إعادة تسمية العمود
            pivot_by_order.rename(columns={col("mrp_controller"): "MRP Contor"}, inplace=True)


            # ترتيب الأعمدة
            cols = pivot_by_order.columns.tolist()
            fixed_order = [col("component"), col("component_desc"), "MRP Contor", col("component_uom"), col("current_stock"), col("component_order_type")]
            other_cols = [c for c in cols if c not in fixed_order]
            pivot_by_order = pivot_by_order[[c for c in fixed_order if c in pivot_by_order.columns] + other_cols]

        else:
            pivot_by_order = pd.DataFrame(columns=[col("component"), col("component_desc"), "MRP Contor", col("component_uom"), col("current_stock"), col("component_order_type")])


# ==============================================================================
        # تحليل الرصيد والمكونات الحرجة مع فلتر MRP Contor ونوع الطلب
# ==============================================================================
        st.markdown("---")
        st.subheader("📊 تحليل حرجية الرصيد ونسبة التغطية")

        # حساب إجمالي الاحتياج والرصيد لكل مكون
        # ✅ التعديل: تم إضافة col("mrp_controller") إلى قائمة التجميع
        component_analysis = merged_df.groupby([
            col("component"), col("component_desc"), col("component_uom"), 
            col("current_stock"), col("component_order_type"), col("hierarchy_level"), col("mrp_controller")
        ]).agg({
            "Required Component Quantity": "sum",
            col("order_type"): lambda x: ", ".join(sorted(set(str(v) for v in x if pd.notna(v))))
        }).reset_index()

        # ❌ تم حذف دمج بيانات MRP Contor
# ==============================================================================
        # استبدال القيم الفارغة بـ "غير محدد"
# ==============================================================================
        # ✅ إعادة تسمية عمود MRP Contor الناتج من التجميع
        component_analysis.rename(columns={col("mrp_controller"): "MRP Contor"}, inplace=True)
        component_analysis["MRP Contor"] = component_analysis["MRP Contor"].fillna("غير محدد")

        # حساب نسبة التغطية
        component_analysis["Coverage Percentage"] = (component_analysis[col("current_stock")] / component_analysis["Required Component Quantity"] * 100).round(1)
        component_analysis["Coverage Status"] = component_analysis["Coverage Percentage"].apply(
            lambda x: "🟢 كافية" if x >= 100 else "🟡 جزئية" if x >= 50 else "🔴 غير كافية"
        )

        # تحديد الأولوية بناء على نسبة التغطية والكمية المطلوبة
        component_analysis["Priority"] = component_analysis.apply(
            lambda row: "🔥 عاجل" if row["Coverage Percentage"] < 30 and row["Required Component Quantity"] > 1000 
            else "⚠️ متوسط" if row["Coverage Percentage"] < 50 
            else "✅ منخفض", 
            axis=1
        )
        # ----- فلاتر المستخدم -----
        mrp_controllers = sorted(component_analysis["MRP Contor"].dropna().unique())
        selected_mrp = st.multiselect("🔍 تصفية حسب MRP Contor:", options=mrp_controllers, default=mrp_controllers, help="اختر واحد أو أكثر من MRP Contor لعرضها")

        component_order_types = sorted(component_analysis[col("component_order_type")].dropna().unique())
        selected_order_types = st.multiselect("🔍 تصفية حسب نوع طلب المكون:", options=component_order_types, default=component_order_types,
            help="اختر نوع طلب المكون (شراء/تصنيع/غير محدد)")

        hierarchy_levels = sorted(component_analysis[col("hierarchy_level")].dropna().unique())
        selected_levels = st.multiselect("🔍 تصفية حسب المستوى الهرمي (Hierarchy Level):", options=hierarchy_levels, default=hierarchy_levels, help="اختر واحد أو أكثر من المستوى لعرضها")

        # تطبيق الفلتر معاً
        filtered_analysis = component_analysis[
            (component_analysis["MRP Contor"].isin(selected_mrp)) &
            (component_analysis[col("component_order_type")].isin(selected_order_types)) &
            (component_analysis[col("hierarchy_level")].isin(selected_levels))
        ]

        # عرض جدول التحليل
        st.dataframe(filtered_analysis.sort_values("Coverage Percentage"))

        # إحصائيات ونسب التغطية بعد التصفية
        total_components = len(filtered_analysis)
        sufficient_coverage = len(filtered_analysis[filtered_analysis["Coverage Percentage"] >= 100])
        partial_coverage = len(filtered_analysis[(filtered_analysis["Coverage Percentage"] >= 50) & (filtered_analysis["Coverage Percentage"] < 100)])
        insufficient_coverage = len(filtered_analysis[filtered_analysis["Coverage Percentage"] < 50])
        critical_components = len(filtered_analysis[filtered_analysis["Priority"] == "🔥 عاجل"])

        st.markdown(f"""
        <div style="direction:rtl; text-align:right; font-size:18px;">
        <span style="font-size:20px; color:#1976d2;">📈 <b>إحصائيات نسبة التغطية</b></span>
        <br><br>
        <ul style="list-style-type:none; padding-right:0;">
            <li>🟢 <b>{sufficient_coverage}</b> مكونات ذات تغطية كافية ({sufficient_coverage/total_components*100:.1f}%)</li>
            <li>🟡 <b>{partial_coverage}</b> مكونات ذات تغطية جزئية ({partial_coverage/total_components*100:.1f}%)</li>
            <li>🔴 <b>{insufficient_coverage}</b> مكونات ذات تغطية غير كافية ({insufficient_coverage/total_components*100:.1f}%)</li>
            <li>🔥 <b style="color:red;">{critical_components}</b> مكونات حرجة تحتاج اهتمام عاجل</li>
        </ul>
        </div>
        """, unsafe_allow_html=True)

        # تحليل إضافي لنوع طلب المكون
        st.markdown("---")
        st.subheader("📊 تحليل المكونات حسب نوع الطلب")

        order_type_stats = filtered_analysis.groupby(col("component_order_type")).agg({
            col("component"): "count",
            "Required Component Quantity": "sum",
            col("current_stock"): "sum"
        }).reset_index()

        order_type_stats["Coverage Percentage"] = (order_type_stats[col("current_stock")] / order_type_stats["Required Component Quantity"] * 100).round(1)

        st.dataframe(order_type_stats)

        # المكونات الحرجة التي تحتاج اهتمام عاجل بعد التصفية
        critical_items = filtered_analysis[filtered_analysis["Priority"] == "🔥 عاجل"]
        if not critical_items.empty:
            st.error("🚨 المكونات الحرجة التي تحتاج إلى اهتمام عاجل:")
            st.dataframe(critical_items[[col("component"), col("component_desc"), "MRP Contor", col("component_order_type"), col("current_stock"), "Required Component Quantity", "Coverage Percentage", "Priority"]])
        else:
            st.success("✅ لا توجد مكونات حرجة تحتاج إلى اهتمام عاجل")

        # رسم بياني لتوزيع نسبة التغطية حسب MRP Contor
        if len(selected_mrp) > 0:
            fig_coverage = px.pie(
                filtered_analysis, 
                names="Coverage Status", 
                title="توزيع المكونات حسب حالة التغطية",
                color="Coverage Status",
                color_discrete_map={"🟢 كافية": "green", "🟡 جزئية": "orange", "🔴 غير كافية": "red"}
            )
            st.plotly_chart(fig_coverage, use_container_width=True)

        # رسم بياني للمكونات الأكثر حرجية مرتبة حسب كمية الطلب
        top_critical = filtered_analysis.nsmallest(10, "Coverage Percentage")
        if not top_critical.empty:
            # تحويل الأعمدة إلى نص قبل الدمج
            top_critical = top_critical.copy()
            top_critical[col("component")] = top_critical[col("component")].astype(str)
            top_critical[col("component_desc")] = top_critical[col("component_desc")].astype(str)
            
            # إنشاء تسمية مختصرة تجمع بين الكود والوصف
            top_critical["Short_Label"] = top_critical[col("component")] + " - " + top_critical[col("component_desc")].str[:20]
            
            # ترتيب المكونات حسب كمية الطلب (من الأكبر إلى الأصغر)
            top_critical = top_critical.sort_values("Required Component Quantity", ascending=True)
            
            fig_critical = px.bar(
                top_critical,
                y="Short_Label",  # التسمية المختصرة على المحور Y
                x="Required Component Quantity",  # كمية الطلب على المحور X
                color="Coverage Percentage",  # التلوين حسب نسبة التغطية
                orientation='h',  # رسم أفقي
                title="أقل 10 مكونات في نسبة التغطية (مرتبة حسب كمية الطلب)",
                labels={
                    "Required Component Quantity": "كمية الطلب المطلوبة", 
                    "Short_Label": "المكون", 
                    "Coverage Percentage": "نسبة التغطية %",
                    "MRP Contor": "MRP Controller"
                },
                hover_data={
                    col("component"): True,
                    col("component_desc"): True,
                    col("current_stock"): True,
                    "Coverage Percentage": ":.1f",
                    "MRP Contor": True,
                    col("component_order_type"): True
                },
                color_continuous_scale="RdYlGn_r"  # مقياس ألوان عكسي (أحمر للأقل تغطية)
            )
            
            # تخصيص التنسيق
            fig_critical.update_traces(
                hovertemplate=(
                    "<b>%{customdata[0]}</b><br>"
                    "الوصف: %{customdata[1]}<br>"
                    "الرصيد الحالي: %{customdata[2]:,}<br>"
                    "الطلب المطلوب: %{x:,}<br>"
                    "نسبة التغطية: %{customdata[3]:.1f}%<br>"
                    "MRP Controller: %{customdata[4]}<br>"
                    "نوع الطلب: %{customdata[5]}"
                )
            )
            
            # تحسين تخطيط الرسم البياني
            fig_critical.update_layout(
                yaxis={'categoryorder':'total ascending'},  # ترتيب حسب القيمة
                xaxis_title="كمية الطلب المطلوبة",
                yaxis_title="المكون",
                hovermode="closest",
                coloraxis_colorbar=dict(title="نسبة التغطية %"),
                height=500  # زيادة الارتفاع لعرض أفضل
            )
            
            # إضافة تسميات القيم على الأعمدة
            fig_critical.update_traces(
                text=top_critical["Required Component Quantity"].apply(lambda x: f"{x:,.0f}"),
                textposition='outside'
            )
            
            st.plotly_chart(fig_critical, use_container_width=True)

        # رسم بياني إضافي لتوزيع المكونات حسب MRP Contor والحالة
        if len(selected_mrp) > 0:
            fig_mrp_coverage = px.sunburst(
                filtered_analysis,
                path=['MRP Contor', 'Coverage Status'],
                values='Required Component Quantity',
                title='توزيع المكونات حسب MRP Contor وحالة التغطية'
            )
            st.plotly_chart(fig_mrp_coverage, use_container_width=True)

        # رسم بياني لتوزيع المكونات حسب نوع الطلب
        fig_order_type = px.pie(
            filtered_analysis, 
            names=col("component_order_type"), 
            title="توزيع المكونات حسب نوع الطلب",
            color=col("component_order_type")
        )
        st.plotly_chart(fig_order_type, use_container_width=True)

# ==============================================================================
        # جدول الكميات الشهرية + الرسم البياني
# ==============================================================================
        if date_cols:
            orders_summary = plan_df.melt(
                id_vars=[col("material"), col("material_desc"), col("order_type")], 
                value_vars=date_cols,
                var_name="Month", 
                value_name="Quantity"
            )
            orders_summary["Month"] = pd.to_datetime(orders_summary["Month"]).dt.month_name()
            orders_grouped = orders_summary.groupby(["Month", col("order_type")]).agg({"Quantity": "sum"}).reset_index()
            pivot_df = orders_grouped.pivot_table(index="Month", columns=col("order_type"), values="Quantity", aggfunc="sum", fill_value=0).reset_index()
            
            if "E" not in pivot_df.columns: pivot_df["E"] = 0
            if "L" not in pivot_df.columns: pivot_df["L"] = 0

            pivot_df["الإجمالي"] = pivot_df["E"] + pivot_df["L"]
            total_sum = pivot_df["الإجمالي"].sum()
            if total_sum > 0:
                pivot_df["E%"] = ((pivot_df["E"] / pivot_df["الإجمالي"]) * 100).round(1).astype(str) + "%"
                pivot_df["L%"] = ((pivot_df["L"] / pivot_df["الإجمالي"]) * 100).round(1).astype(str) + "%"
            else:
                 pivot_df["E%"], pivot_df["L%"] = "0.0%", "0.0%"

            month_order = {m: i for i, m in enumerate(calendar.month_name) if m}
            pivot_df = pivot_df.sort_values(by="Month", key=lambda x: x.map(month_order))

            st.subheader("📊 توزيع الكميات الشهرية حسب نوع الأمر")
            html_table = "<table border='1' style='border-collapse: collapse; width:100%; text-align:center; color:green;'>"
            html_table += "<tr style='background-color:#4CAF50; color:white;'><th>الشهر</th><th>E</th><th>L</th><th>الإجمالي</th><th>E%</th><th>L%</th></tr>"
            for _, row in pivot_df.iterrows():
                html_table += "<tr>"
                html_table += f"<td style='color:blue; font-weight:bold;'>{row['Month']}</td><td>{int(row.get('E', 0))}</td><td>{int(row.get('L', 0))}</td><td>{int(row.get('الإجمالي', 0))}</td><td>{row.get('E%', '')}</td><td>{row.get('L%', '')}</td>"
                html_table += "</tr>"
            html_table += "</table>"
            st.markdown(f"<div style='direction:rtl;'>{html_table}</div>", unsafe_allow_html=True)

            # تحسين الرسم البياني بإضافة تسميات عربية
            fig = px.bar(
                pivot_df, 
                x="Month", 
                y=["E", "L"], 
                barmode="group", 
                text_auto=True, 
                title="رسم بياني لتوزيع الكميات",
                labels={"value": "الكمية", "variable": "نوع الأمر", "Month": "الشهر"},
                template="streamlit"
            )
            st.plotly_chart(fig, use_container_width=True)
            st.markdown("---")

# ==============================================================================
        # تحويل رؤوس الأعمدة التي تحتوي على تواريخ إلى صيغة مختصرة "يوم شهر"
# ==============================================================================
        plan_df.columns = [
            col.strftime("%d %b") if isinstance(col, (datetime.datetime, pd.Timestamp)) else col
            for col in plan_df.columns
        ]
# ==============================================================================
        # 📆 تحليل الطلب الشهري للمكونات (الخامات MET فقط)
# ==============================================================================
        st.subheader("📆 تحليل الطلب الشهري للمكونات (الخامات MET فقط)")

        # 🔹 فلترة المكونات الخام (التي تبدأ برقم 1) وMRP Contor = MET فقط
        raw_materials_df = merged_df[
            merged_df[col("component")].astype(str).str.startswith("1")
        ].copy()

        # ✅ الآن نعتمد على عمود MRP Controller الموجود في merged_df
        raw_materials_df.rename(columns={col("mrp_controller"): "MRP Contor"}, inplace=True)
        raw_materials_df = raw_materials_df[
            raw_materials_df["MRP Contor"].fillna("") == "MET"
        ]


# 🔹 توحيد وحدات الوزن: (تم تطبيقها مسبقاً في دالة load_and_validate_data)
        # ملاحظة: تم توحيد الوحدات (جرام -> كجم) في دالة load_and_validate_data.
        # لذلك، نحتاج فقط لإعادة تسمية عمود الكمية للوضوح.
        
        raw_materials_df.rename(
            columns={"Required Component Quantity": "Required Component Quantity (KG)"},
            inplace=True
        )
        
        # 🔹 نضمن أن الوحدة هي KG في هذا الجدول لتبسيط العرض (حيث تم التحويل مسبقاً)
        raw_materials_df[col("component_uom")] = "KG"


        # 🔹 تجميع القيم حسب الشهر والمكون
        monthly_raw = raw_materials_df.groupby(
            [col("component"), col("component_desc"), col("component_uom"), "Date"]
        )["Required Component Quantity (KG)"].sum().reset_index()

        # 🔹 Pivot بالشهر
        pivot_raw_monthly = monthly_raw.pivot_table(
            index=[col("component"), col("component_desc"), col("component_uom")],
            columns="Date",
            values="Required Component Quantity (KG)",
            aggfunc="sum",
            fill_value=0
        ).reset_index()

        # 🔹 تنسيق أعمدة التاريخ لتظهر بشكل واضح (مثلاً: 01 Nov)
        pivot_raw_monthly.columns = [
            col.strftime("%d %b") if isinstance(col, pd.Timestamp) else col
            for col in pivot_raw_monthly.columns
        ]

        # 🔹 عرض النتائج في الواجهة
        st.dataframe(pivot_raw_monthly, use_container_width=True)

        # 🔹 إنشاء زر التحميل الفوري بعد إنشاء الملف
        if not pivot_raw_monthly.empty:
            raw_excel_buffer = BytesIO()
            with pd.ExcelWriter(raw_excel_buffer, engine="openpyxl") as writer:
                pivot_raw_monthly.to_excel(writer, sheet_name="Raw_Materials_MET", index=False)
            raw_excel_buffer.seek(0)

            st.download_button(
                label="📥 تحميل ملف تحليل الخامات (MET)",
                data=raw_excel_buffer,
                file_name=f"Raw_Materials_Analysis_MET_{datetime.datetime.now().strftime('%d_%b_%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("✅(MET)تم إنشاء الملف الخاص بخامات المعادن فقط  بنجاح وجاهز للتحميل اضغط اعلاه  تحميل ملف تحليل الخامات .")
# ==============================================================================
        # 5. حساب جدول (Component in BOMs) - المكونات داخل الموديلات
# ==============================================================================
        st.subheader("📋 قائمة الموديلات التي تستخدم كل مكون")

        # 1. دمج بيانات الخطة (Plan) مع بيانات المكونات الأصلية (Component)
        # نحتاج إلى (material, Planned Quantity, Order Type) من جدول الخطة
        
        # تحويل الخطة من الشكل الأفقي إلى الطولي لجمع إجمالي الكمية المطلوبة لكل موديل
        plan_summary = plan_melted.groupby(
            [col("material"), col("order_type")]
        )["Planned Quantity"].sum().reset_index()
        
        # إعادة تسمية عمود الكمية
        plan_summary.rename(columns={"Planned Quantity": "plan_qty"}, inplace=True)

        # دمج الكميات ونوع الطلب إلى جدول المكونات الأصلي
        component_bom_merged = pd.merge(
            component_df_orig, 
            plan_summary, 
            on=col("material"), 
            how="left"
        ).fillna({"plan_qty": 0, col("order_type"): 'N/A'})


        # 2. إنشاء عمود تجميعي (model_info) للمحور الأفقي
        # (الموديل + كمية الخطة + نوع الطلب)
        component_bom_merged["model_info"] = (
            component_bom_merged[col("material")].astype(str)
            + " ("
            + component_bom_merged["plan_qty"].astype(int).astype(str) # تحويل لرقم صحيح
            + " "
            + component_bom_merged[col("order_type")].astype(str)
            + ")"
        )

        # 3. إنشاء جدول محوري يضم كمية المكون + معلومات الخطة + نوع الطلب
        component_bom_pivot = component_bom_merged.pivot_table(
            index=[
                col("component"),
                col("component_desc"),
                col("mrp_controller"),
                col("component_uom")
            ],
            columns="model_info",
            values=col("component_qty"),
            aggfunc="first",
            fill_value=0
        ).reset_index()

        # إعادة تسمية عمود MRP Contor
        component_bom_pivot.rename(columns={col("mrp_controller"): "MRP Contor"}, inplace=True)
        
        # 4. عرض الجدول في الواجهة
        st.dataframe(component_bom_pivot, use_container_width=True)
        st.markdown("---")
# ==============================================================================
        # زر إنشاء النسخة الكاملة (التصدير)
# ==============================================================================
        if st.button("🗜️ اضغط هنا لإنشاء النسخة الكاملة"):
            with st.spinner('⏳ جاري إنشاء الملفات وتجهيزها للتحميل...'):
                current_date = datetime.datetime.now().strftime("%d_%b_%Y")
                excel_buffer = BytesIO()
           
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:

                    # -------------------------------
                    # شيت الخطة الأساسية
                    # -------------------------------
                    try:
                        plan_df.to_excel(writer, sheet_name="Plan", index=False)
                    except:
                        pd.DataFrame().to_excel(writer, sheet_name="Plan", index=False)

                    # -------------------------------
                    # شيت الملخص
                    # -------------------------------
                    try:
                        summary_df.to_excel(writer, sheet_name="Summary", index=False)
                    except:
                        pd.DataFrame().to_excel(writer, sheet_name="Summary", index=False)

                    # -------------------------------
                    # شيت نتائج الـ Recursive BOM
                    # -------------------------------
                    try:
                        pivot_recursive.to_excel(writer, sheet_name="Recursive_BOM_Results", index=False)
                    except:
                        try:
                            agg_recursive.to_excel(writer, sheet_name="Recursive_BOM_Results", index=False)
                        except:
                            pd.DataFrame().to_excel(writer, sheet_name="Recursive_BOM_Results", index=False)

                    # -------------------------------
                    # شيت الاحتياجات حسب التاريخ
                    # -------------------------------
                    try:
                        pivot_by_date.to_excel(writer, sheet_name="Need_By_Date", index=False)
                    except:
                        pd.DataFrame().to_excel(writer, sheet_name="Need_By_Date", index=False)

                    # -------------------------------
                    # شيت الاحتياجات حسب نوع الطلب
                    # -------------------------------
                    try:
                        pivot_by_order.to_excel(writer, sheet_name="Need_By_Order Type", index=False)
                    except:
                        pd.DataFrame().to_excel(writer, sheet_name="Need_By_Order Type", index=False)

                    # -------------------------------
                    # شيت تحليل تغطية المخزون للمكونات
                    # -------------------------------
                    try:
                        component_analysis.to_excel(writer, sheet_name="Stock_Coverage_Analysis", index=False)
                    except:
                        pd.DataFrame().to_excel(writer, sheet_name="Stock_Coverage_Analysis", index=False)

                    # -------------------------------
                    # شيت المكونات في BOMs
                    # -------------------------------
                    try:
                        component_bom_pivot.reset_index().to_excel(writer, sheet_name="Component_in_BOMs", index=False)
                    except:
                        pd.DataFrame().to_excel(writer, sheet_name="Component_in_BOMs", index=False)

                    # -------------------------------
                    # شيت المكونات الأساسي
                    # -------------------------------
                    try:
                        component_df.to_excel(writer, sheet_name="Component", index=False)
                    except:
                        pd.DataFrame().to_excel(writer, sheet_name="Component", index=False)

                    # -------------------------------
                    # شيت MRP Contor إذا كان موجود
                    # -------------------------------
                    try:
                        if not mrp_df.empty:
                            mrp_df.to_excel(writer, sheet_name="MRP Contor", index=False)
                    except:
                        pass

                excel_buffer.seek(0)
# ==============================================================================
                # زر التحميل
# ==============================================================================
                st.subheader("🔥 أضغط هنا لتحميل النسخة الإكسل الكاملة ")
                st.download_button(
                    label=" 📊  تحميل ملف الإكسل",
                    data=excel_buffer, 
                    file_name=f"All_Component_Results_{current_date}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.balloons()
                st.success("✅ تم إنشاء النسخة الكاملة بنجاح، وجميع الشيتات موجودة داخل Excel")
# ==============================================================================
# --- التذييل ---
# ==============================================================================
st.markdown(
    """
    <hr>
    <div style="text-align:center; direction:rtl; font-size:14px; color:gray;">
        ✨ تم التنفيذ بواسطة <b>م / رضا رشدي</b> – جميع الحقوق محفوظة © 2025 ✨
    </div>
    """,
    unsafe_allow_html=True
)
