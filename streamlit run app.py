# =====================("📂 MRP  برنامج تحليل واستخراج وحفظ نتائج الـ")======================
# MRP Analysis Tool - Multi-Level BOM Explosion
# Developed by: Reda Roshdy
# Fixed & Enhanced by: Claude (Anthropic)
# Date: Mar-2026
# =======================================================================

# -------------------------------
# 1. استدعاء المكتبات اللازمة
# -------------------------------
import streamlit as st
import pandas as pd
import datetime
from io import BytesIO
import calendar
import plotly.express as px

# ==============================================================================
# 2. إعداد التكوين والأعمدة
# ==============================================================================
COLUMN_NAMES = {
    "material":             ["Material", "Item", "code", "Code", "المادة", "Product"],
    "material_desc":        ["Material Description", "Description", "وصف"],
    "order_type":           ["Order Type", "OT", "نوع الطلب", "Sales Org."],
    "component":            ["Component", "Comp", "المكون"],
    "component_desc":       ["Component Description", "Comp Desc", " المسمى", "وصف المكون"],
    "component_uom":        ["Component UoM", "UoM", "الوحدة"],
    "component_qty":        ["Component Quantity", "Qty", "كمية المكون"],
    "base_qty":             ["Base Quantity", "Base Qty", "الكمية الأساسية"],
    "mrp_controller":       ["MRP Controller", "مسؤول MRP"],
    "current_stock":        ["Current Stock", "Stock", "المخزون الحالي", "Unrestricted"],
    "component_order_type": ["Component Order Type", "Order Category", "نوع أمر المكون", "Procurement Type"],
    "hierarchy_level":      ["Hierarchy Level", "Level", "المستوى الهرمي"],
    "parent_material":      ["Parent Material", "Direct Parent", "الأب المباشر"],
}

def col(name_key):
    """إرجاع اسم العمود الرئيسي"""
    return COLUMN_NAMES[name_key][0]

def normalize_columns(df, column_map):
    """توحيد أسماء الأعمدة إلى الاسم الرئيسي"""
    rename_dict = {}
    for key, aliases in column_map.items():
        for alias in aliases:
            if alias in df.columns and alias != aliases[0]:
                rename_dict[alias] = aliases[0]
    return df.rename(columns=rename_dict)

# ==============================================================================
# 3. دالة تحميل البيانات والتحقق منها
# ==============================================================================
@st.cache_data
def load_and_validate_data(uploaded_file):
    try:
        xls = pd.ExcelFile(uploaded_file, engine='openpyxl')

        # --- التحقق من الأوراق ---
        required_sheets = ["plan", "Component"]
        missing_sheets = [s for s in required_sheets if s not in xls.sheet_names]
        if missing_sheets:
            st.error(f"❌ الملف لا يحتوي على الأوراق المطلوبة: {', '.join(missing_sheets)}")
            st.stop()

        # --- تحميل البيانات ---
        plan_df      = normalize_columns(xls.parse("plan"),      COLUMN_NAMES)
        component_df = normalize_columns(xls.parse("Component"), COLUMN_NAMES)
        mrp_df = (
            normalize_columns(xls.parse("MRP Controller"), COLUMN_NAMES)
            if "MRP Controller" in xls.sheet_names
            else pd.DataFrame()
        )

        # ✅ إزالة الأعمدة الزائدة غير المعروفة من ورقة Component
        known_cols = [aliases[0] for aliases in COLUMN_NAMES.values()]
        extra_cols = [c for c in component_df.columns if c not in known_cols]
        if extra_cols:
            component_df.drop(columns=extra_cols, inplace=True)

        # --- التحقق من الأعمدة الأساسية ---
        required_plan_cols = [col("material"), col("material_desc"), col("order_type")]
        if not all(c in plan_df.columns for c in required_plan_cols):
            st.error(f"❌ جدول الخطة ناقص أعمدة: {required_plan_cols}")
            st.stop()

        required_comp_cols = [col("material"), col("component"), col("component_qty")]
        if not all(c in component_df.columns for c in required_comp_cols):
            st.error(f"❌ جدول المكونات ناقص أعمدة: {required_comp_cols}")
            st.stop()

        # --- تنظيف الأعمدة الرقمية في component_df ---
        comp_qty_col = col("component_qty")
        base_qty_col = col("base_qty")

        component_df[comp_qty_col] = pd.to_numeric(component_df[comp_qty_col], errors='coerce').fillna(0)

        # ✅ FIX 1: تطبيق Base Qty بشكل صحيح (خارج except)
        if base_qty_col in component_df.columns:
            component_df[base_qty_col] = (
                pd.to_numeric(component_df[base_qty_col], errors='coerce')
                .fillna(1).replace(0, 1)
            )
            # ⚠️ تحذير عند وجود أصفار في Base Quantity
            zero_base = (pd.to_numeric(xls.parse("Component").get(base_qty_col, pd.Series(dtype=float)),
                                       errors='coerce') == 0).sum() if base_qty_col in xls.parse("Component").columns else 0
            if zero_base > 0:
                st.warning(f"⚠️ يوجد {zero_base} قيمة صفرية في عمود Base Quantity — تم استبدالها بـ 1 تلقائياً. تحقق من البيانات.")
            component_df[comp_qty_col] = component_df[comp_qty_col] / component_df[base_qty_col]
            component_df.drop(columns=[base_qty_col], inplace=True)

        # --- الأعمدة الاختيارية مع قيم افتراضية ---
        if col("current_stock") not in component_df.columns:
            component_df[col("current_stock")] = 0
        else:
            component_df[col("current_stock")] = pd.to_numeric(
                component_df[col("current_stock")], errors='coerce'
            ).fillna(0)

        if col("component_order_type") not in component_df.columns:
            component_df[col("component_order_type")] = "غير محدد"

        if col("hierarchy_level") not in component_df.columns:
            component_df[col("hierarchy_level")] = 1
        else:
            component_df[col("hierarchy_level")] = pd.to_numeric(
                component_df[col("hierarchy_level")], errors='coerce'
            ).fillna(1).astype(int)

        if col("component_desc") not in component_df.columns:
            component_df[col("component_desc")] = ""

        if col("component_uom") not in component_df.columns:
            component_df[col("component_uom")] = ""

        if col("mrp_controller") not in component_df.columns:
            component_df[col("mrp_controller")] = "غير محدد"

        # ✅ تنظيف عمود Parent Material إن وُجد
        # هذا العمود يحتوي على الأب المباشر الفعلي لكل مكون (من SAP CS12)
        if col("parent_material") in component_df.columns:
            component_df[col("parent_material")] = (
                component_df[col("parent_material")].astype(str).str.strip()
            )
        # إذا لم يكن موجوداً → نُنشئه من Material (fallback للتوافق مع ملفات قديمة)
        else:
            component_df[col("parent_material")] = component_df[col("material")]

        # ✅ توحيد وحدات الوزن إلى كيلوجرام
        # أي مكون وحدته G أو g أو GM أو gram → نقسم الكمية والرصيد على 1000 ونغير الوحدة إلى KG
        gram_variants = {"g", "gm", "gr", "gram", "grams", "جرام", "جم"}
        uom_col = col("component_uom")
        qty_col = col("component_qty")
        stk_col = col("current_stock")

        is_gram = component_df[uom_col].astype(str).str.strip().str.lower().isin(gram_variants)

        if is_gram.any():
            component_df.loc[is_gram, qty_col] = component_df.loc[is_gram, qty_col] / 1000
            component_df.loc[is_gram, stk_col] = component_df.loc[is_gram, stk_col] / 1000
            component_df.loc[is_gram, uom_col] = "KG"

        # ✅ NEW: توحيد وحدات المساحة من CM2 إلى M2
        cm2_variants = {"cm2", "cm^2", "cm²", "سم2", "سم²"}
        is_cm2 = component_df[uom_col].astype(str).str.strip().str.lower().isin(cm2_variants)

        if is_cm2.any():
            component_df.loc[is_cm2, qty_col] = component_df.loc[is_cm2, qty_col] / 10000
            component_df.loc[is_cm2, stk_col] = component_df.loc[is_cm2, stk_col] / 10000
            component_df.loc[is_cm2, uom_col] = "M2"

        return plan_df, component_df, mrp_df

    except Exception as e:
        st.error(f"❌ فشل تحميل الملف: {str(e)}")
        st.stop()

# ==============================================================================
# ✅ FIX 2: دالة BOM Explosion متعددة المستويات (الإصلاح الجوهري)
# ==============================================================================
def bom_explosion(plan_melted, component_df):
    """
    Multi-Level BOM Explosion — النهج الصحيح لـ SAP CS12

    المفتاح الذهبي:  Material + Parent Material + Component
    ✔ يمنع دمج نفس (Parent→Component) من منتجات مختلفة
    ✔ يجمع الكميات داخل نفس المنتج فقط
    ✔ يعزل bom_dict لكل Material → explosion آمن بدون تلوث

    الخوارزمية:
    1. groupby(Material + Parent + Component) → bom_core فريد لكل منتج
    2. bom_dict per Material → tree[parent] = [(comp, qty), ...]
    3. explode تعاودي لكل صف في الخطة مستقلاً
    """
    from collections import defaultdict

    component_df = component_df.copy()
    component_df[col("component")] = component_df[col("component")].astype(str).str.strip()
    component_df[col("material")]  = component_df[col("material")].astype(str).str.strip()

    # تحديد عمود الأب المباشر
    has_parent_col = col("parent_material") in component_df.columns
    if has_parent_col:
        component_df[col("parent_material")] = component_df[col("parent_material")].astype(str).str.strip()
        parent_col = col("parent_material")
    else:
        parent_col = col("material")

    # ✅ STEP 1: تنظيف ثم groupby(Material + Parent + Component) + sum
    #
    # SAP CS12 يصدر أحياناً صفوفاً مكررة حرفياً لنفس الزوج (Parent→Component)
    # بنفس الكمية — هذه نسخ وليست كميات إضافية حقيقية.
    # الحل الصحيح: خطوتان:
    #   1. drop_duplicates على كل الأعمدة → يحذف النسخ الحرفية
    #   2. groupby(Material+Parent+Component)+sum → يجمع الكميات الحقيقية المختلفة
    #      داخل نفس المنتج، مع عزل كامل بين المنتجات المختلفة
    component_df = component_df.drop_duplicates(
        subset=[col("material"), parent_col, col("component"), col("component_qty")],
        keep="first"
    )
    bom_core = component_df.groupby(
        [col("material"), parent_col, col("component")],
        as_index=False
    )[col("component_qty")].sum()

    # ✅ STEP 2: بناء bom_dict منفصل لكل Material
    # tree[material][parent] = [(component, qty), ...]
    bom_dict = {}
    for mat, group in bom_core.groupby(col("material")):
        tree = defaultdict(list)
        for _, row in group.iterrows():
            tree[row[parent_col]].append(
                (row[col("component")], row[col("component_qty")])
            )
        bom_dict[mat] = tree

    # معلومات وصفية للمكونات
    comp_info = (
        component_df
        .drop_duplicates(subset=[col("component")], keep="last")
        .set_index(col("component"))[[
            col("component_desc"),
            col("component_uom"),
            col("mrp_controller"),
            col("current_stock"),
            col("component_order_type"),
        ]]
    )

    # ✅ STEP 3: دالة explosion تعاودية آمنة
    def explode(root_material, parent, qty, path, level, row_buf):
        """
        root_material : المنتج الجذر (لجلب bom_dict الصحيح)
        parent        : الأب الحالي الذي نبحث عن أبنائه
        qty           : الكمية المطلوبة من الأب الحالي
        path          : مسار العقد التي مررنا بها (لمنع الحلقات)
        level         : المستوى الهرمي الحالي
        row_buf       : مخزن الصفوف الناتجة

        المنطق الصحيح لحساب الكميات:
        - نبحث أولاً في bom_dict[root_material] عن أبناء parent
        - إذا لم نجد (مكون وسيط له BOM مستقل)، نبحث في bom_dict[parent]
        - الكمية المطلوبة = qty_from_parent × qty_of_this_child
        """
        if parent in path or level > 10:
            return

        # البحث في شجرة المنتج الجذر أولاً، ثم في شجرة الأب نفسه (نصف مصنّع)
        tree = bom_dict.get(root_material, {})
        children = tree.get(parent, [])

        if not children:
            # المكون الوسيط قد يكون له BOM مستقل (semi-finished product)
            tree = bom_dict.get(parent, {})
            children = tree.get(parent, [])

        if not children:
            return

        new_path = path | {parent}
        for comp, comp_qty in children:
            # ✅ الكمية الصحيحة: كمية الأب × كمية المكون لكل وحدة من الأب
            needed = qty * comp_qty
            row_buf.append({
                "Parent":                        parent,
                col("component"):                comp,
                col("component_qty"):            comp_qty,
                "Required Component Quantity":   needed,
                "BOM Level":                     level,
            })
            # 🔁 الاستدعاء العودي الصحيح:
            # - نمرر comp كـ parent الجديد (الأب للمستوى التالي)
            # - نمرر needed كـ qty (الكمية المطلوبة من comp)
            # - نحاول أولاً داخل شجرة root_material، وإلا داخل شجرة comp نفسه
            explode(root_material, comp, needed, new_path, level + 1, row_buf)

    # ✅ STEP 4: تشغيل الـ explosion لكل صف في الخطة
    all_rows = []
    for _, plan_row in plan_melted[plan_melted["Planned Quantity"] > 0].iterrows():
        mat  = str(plan_row[col("material")]).strip()
        qty  = plan_row["Planned Quantity"]
        ot   = plan_row[col("order_type")]
        date = plan_row["Date"]

        row_buf = []
        explode(mat, mat, qty, set(), level=1, row_buf=row_buf)
        mat_desc = str(plan_row.get(col("material_desc"), "")).strip()
        for r in row_buf:
            r[col("material")]      = mat
            r[col("material_desc")] = mat_desc
            r["Order Type"]         = ot
            r["Date"]               = date
        all_rows.extend(row_buf)

    if not all_rows:
        return pd.DataFrame()

    result = pd.DataFrame(all_rows)

    # إضافة الأعمدة الوصفية
    comp_info_clean = (
        comp_info.reset_index()
        .rename(columns={col("component"): "_comp_key"})
        .drop_duplicates(subset=["_comp_key"])
    )
    result = result.merge(
        comp_info_clean,
        left_on=col("component"),
        right_on="_comp_key",
        how="left"
    ).drop(columns=["_comp_key"], errors="ignore")

    return result

# ==============================================================================
# 4. واجهة المستخدم
# ==============================================================================
st.set_page_config(page_title="💪🔥 MRP Tool", page_icon="👍", layout="wide")
st.header("🔥 برنامج تحليل واستخراج وحفظ نتائج الـ MRP 💪🔥")

with st.expander("📖 دليل الاستخدام"):
    st.markdown("""
    ### كيفية استخدام البرنامج:
    - حمّل ملف **Excel** يحتوي على أوراق **(plan و Component و MRP Controller)**
    - احفظ وصدّر النتائج بصيغة Excel

    #### 📋 الأعمدة الأساسية:
    **ورقة plan:**
    - `Material` — كود المنتج
    - `Material Description` — وصف المنتج
    - `Order Type` — E (تصدير) أو L (محلي)
    - أعمدة التواريخ — الكميات المخططة

    **ورقة Component:**
    - `Material` — كود الجذر النهائي (Root)
    - `Parent Material` — **الأب المباشر الفعلي** لكل مكون *(العمود الأساسي للحساب الصحيح)*
    - `Component` — كود المكون
    - `Component Quantity` — الكمية لكل وحدة من الأب المباشر
    - `Base Quantity` *(اختياري)* — الكمية الأساسية للقسمة
    - `Hierarchy Level` — المستوى الهرمي (1، 2، 3، ...)
    - `Current Stock` — الرصيد الحالي
    - `Component Order Type` — F (شراء) أو E (تصنيع)
    """)

st.markdown("<p style='font-size:16px; font-weight:bold;'>📂 اختر ملف الخطة الشهرية Excel</p>", unsafe_allow_html=True)
uploaded_file = st.file_uploader("", type=["xlsx"])

if not uploaded_file:
    st.stop()

# --- تحميل البيانات ---
plan_df, component_df, mrp_df = load_and_validate_data(uploaded_file)
plan_df_orig      = plan_df.copy()
component_df_orig = component_df.copy()
mrp_df_orig       = mrp_df.copy()

# --- استخراج أعمدة التواريخ ---
date_cols = [c for c in plan_df.columns if isinstance(c, (datetime.datetime, pd.Timestamp))]

if not date_cols:
    st.error("❌ لم يتم العثور على أعمدة تواريخ في ورقة الخطة.")
    st.stop()

with st.spinner("⏳ جاري معالجة البيانات..."):

    # ==============================================================================
    # A. تجهيز الخطة (Melt)
    # ==============================================================================
    plan_melted = plan_df.melt(
        id_vars=[col("material"), col("material_desc"), col("order_type")],
        value_vars=date_cols,
        var_name="Date",
        value_name="Planned Quantity"
    )
    plan_melted["Date"] = pd.to_datetime(plan_melted["Date"], errors='coerce')
    plan_melted["Planned Quantity"] = pd.to_numeric(
        plan_melted["Planned Quantity"], errors="coerce"
    ).fillna(0)
    # إزالة الصفوف بكمية صفر أو تاريخ مجهول
    plan_melted = plan_melted[
        (plan_melted["Planned Quantity"] > 0) &
        (plan_melted["Date"].notna())
    ].copy()

    # ==============================================================================
    # ✅ B. تشغيل Multi-Level BOM Explosion
    # ==============================================================================
#    st.markdown("---")
#    st.subheader("🔩 نتائج BOM Explosion — جميع المستويات الهرمية")

    result_df = bom_explosion(plan_melted, component_df)

    if result_df.empty:
        st.warning("⚠️ لم يتم العثور على مكونات مطابقة بين الخطة والـ BOM.")
    else:
        # تجميع إجمالي لكل مكون × تاريخ × نوع الطلب
        # ✅ نعتمد على BOM Level (المحسوب تعاودياً) وليس hierarchy_level من ورقة Component
        merged_df = (
            result_df
            .groupby([
                col("component"),
                col("component_desc"),
                col("component_uom"),
                col("mrp_controller"),
                col("current_stock"),
                col("component_order_type"),
                "Order Type",
                "Date",
                "BOM Level",
            ], as_index=False)
            ["Required Component Quantity"]
            .sum()
        )

        actual_levels = sorted(merged_df["BOM Level"].unique())
#        st.success(
 #           f"✅ إجمالي صفوف الاحتياج: {len(merged_df):,} | "
  #          f"مكونات فريدة: {merged_df[col('component')].nunique():,} | "
   #         f"المستويات المحسوبة: {actual_levels}"
    #    )

        # 🔍 DEBUG: مساعدة في التشخيص — يمكن إخفاؤه بعد التحقق
        with st.expander("🔍 تشخيص: عيّنة من نتائج result_df الخام (قبل التجميع)"):
            debug_sample = result_df[["Parent", col("component"), "Order Type", "Date",
                                      col("component_qty"), "Required Component Quantity", "BOM Level"]].copy()
            debug_sample["Date"] = debug_sample["Date"].astype(str)
            st.dataframe(debug_sample.sort_values(["BOM Level", "Parent", col("component")]).head(100),
                         use_container_width=True)
            st.caption(f"إجمالي الصفوف الخام: {len(result_df):,}")

        # عرض مبسط بالمستوى
        display_cols = [
            col("component"), col("component_desc"),
            col("mrp_controller"), col("component_order_type"),
            "Order Type", "Date", "Required Component Quantity", "BOM Level"
        ]
        display_cols = [c for c in display_cols if c in merged_df.columns]
#        st.dataframe(merged_df[display_cols].sort_values(
 #           ["BOM Level", col("component"), "Date"]
  #      ), use_container_width=True)

    # ==============================================================================
    # C. الملخص السريع
    # ==============================================================================
    st.markdown("---")
    total_models     = plan_df[col("material")].nunique()
    total_components = component_df[col("component")].nunique()
    total_boms       = len(component_df)
    empty_mrp_count  = mrp_df[col("component")].isna().sum() if not mrp_df.empty else 0

    diff_uom = component_df.groupby(col("component"))[col("component_uom")].nunique()
    diff_uom = diff_uom[diff_uom > 1]
    total_diff_uom = len(diff_uom)

    # اضافة المسمى جانب الكود لاكثر من وحدة
#    diff_uom_str   = ", ".join(map(str, diff_uom.index)) if total_diff_uom > 0 else "لا يوجد"
    diff_uom_str = ", ".join(
        f"{comp_code} ({component_df.loc[component_df[col('component')] == comp_code, 'Component Description'].iloc[0]})"
        for comp_code in diff_uom.index) if total_diff_uom > 0 else "لا يوجد"

    diff_uom_color = "red" if total_diff_uom > 0 else "green"

    missing_boms      = set(plan_df[col("material")]) - set(component_df[col("material")])
    total_missing_boms = len(missing_boms)
    missing_boms_html  = (
        f"<span style='color:red;'>{', '.join(map(str, missing_boms))}</span>"
        if missing_boms else "<span style='color:green;'>لا يوجد</span>"
    )

    # إحصائيات نوع الطلب
    order_type_map = {"F": "شراء", "E": "تصنيع"}
    component_df["Order_Type_Label"] = component_df[col("component_order_type")].map(order_type_map).fillna("غير محدد")
    purchase_count      = component_df.loc[component_df["Order_Type_Label"] == "شراء",    col("component")].nunique()
    manufacturing_count = component_df.loc[component_df["Order_Type_Label"] == "تصنيع",   col("component")].nunique()
    undefined_count     = component_df.loc[component_df["Order_Type_Label"] == "غير محدد", col("component")].nunique()

    # المستويات الهرمية الموجودة
    levels_summary = (
        component_df.groupby(col("hierarchy_level"))[col("component")]
        .nunique()
        .reset_index()
        .rename(columns={col("component"): "عدد المكونات", col("hierarchy_level"): "المستوى"})
    )

    st.markdown(f"""
    <div style="direction:rtl; text-align:right; font-size:18px;">
    <span style="font-size:20px; color:#1976d2;">📌 <b>ملخص نتائج الخطة</b></span><br><br>
    <ul style="list-style-type:none; padding-right:0;">
      <li>🟢 <b>{total_models}</b> موديلات بالخطة</li>
      <li>🔵 <b>{total_components}</b> مكون فريد</li>
      <li>🟠 <b>{total_boms}</b> إجمالي سطور الـ BOM</li>
      <li>{"❌" if empty_mrp_count>0 else "✅"} <b style="color:{'red' if empty_mrp_count>0 else 'green'};">{empty_mrp_count}</b> مكونات بدون MRP Controller</li>
      <li>{"⚠️" if total_diff_uom>0 else "✅"} <b style="color:{'red' if total_diff_uom>0 else 'green'};">{total_diff_uom}</b> مكونات لها أكثر من وحدة: <span style="color:{diff_uom_color};">{diff_uom_str}</span></li>
      <li>{"⚠️" if total_missing_boms>0 else "✅"} <b style="color:{'red' if total_missing_boms>0 else 'green'};">{total_missing_boms}</b> منتجات بالخطة بدون BOM: {missing_boms_html}</li>
    </ul>
    </div>
    """, unsafe_allow_html=True)

    st.markdown(f"""
    <div style="direction:rtl; text-align:right; font-size:18px;">
    <span style="font-size:20px; color:#1976d2;">🔹 <b>ملخص أنواع طلب المكونات</b></span><br><br>
    <ul style="list-style-type:none; padding-right:0;">
        <li>🛒 <b>{purchase_count}</b> مكونات شراء (F)</li>
        <li>🏭 <b>{manufacturing_count}</b> مكونات تصنيع (E)</li>
        <li>❓ <b>{undefined_count}</b> مكونات غير محددة</li>
    </ul>
    </div>
    """, unsafe_allow_html=True)

    st.subheader("📊 توزيع المكونات على المستويات الهرمية")
    st.dataframe(levels_summary, use_container_width=True,hide_index=True)

    # ==============================================================================
    # D. Need_By_Date — الاحتياج حسب التاريخ
    # ==============================================================================
    st.markdown("---")
    st.subheader("📅 Need by Date — الاحتياج الكلي لكل مكون حسب التاريخ")

    if not result_df.empty:
        # تجميع كل المستويات: لكل مكون × تاريخ → جمع الاحتياجات
        result_date = (
            merged_df
            .groupby([
                col("component"), col("component_desc"), col("component_uom"),
                col("mrp_controller"), col("current_stock"), col("component_order_type"),
                "Date"
            ], as_index=False)
            ["Required Component Quantity"]
            .sum()
        )

        pivot_by_date = result_date.pivot_table(
            index=[
                col("component"), col("component_desc"), col("component_uom"),
                col("mrp_controller"), col("current_stock"), col("component_order_type"),
            ],
            columns="Date",
            values="Required Component Quantity",
            aggfunc="sum",
            fill_value=0
        ).reset_index()

        # تنسيق أسماء أعمدة التواريخ
        pivot_by_date.columns = [
            c.strftime("%d %b") if isinstance(c, (pd.Timestamp, datetime.datetime)) else c
            for c in pivot_by_date.columns
        ]

        st.dataframe(pivot_by_date, use_container_width=True,hide_index=True)

    # ==============================================================================
    # E. Need_By_Order_Type — الاحتياج حسب التاريخ ونوع الطلب (E / L)
    # ==============================================================================
    st.markdown("---")
    st.subheader("📦 Need by Order Type — الاحتياج مقسّم حسب نوع الطلب والتاريخ")

    if not result_df.empty:
        result_order = (
            merged_df
            .groupby([
                col("component"), col("component_desc"), col("component_uom"),
                col("mrp_controller"), col("current_stock"), col("component_order_type"),
                "Order Type", "Date"
            ], as_index=False)
            ["Required Component Quantity"]
            .sum()
        )

        pivot_by_order = result_order.pivot_table(
            index=[
                col("component"), col("component_desc"), col("component_uom"),
                col("mrp_controller"), col("current_stock"), col("component_order_type"),
            ],
            columns=["Date", "Order Type"],
            values="Required Component Quantity",
            aggfunc="sum",
            fill_value=0
        ).reset_index()

        # تسطيح أسماء الأعمدة المركبة
        flat_cols = []
        for c in pivot_by_order.columns:
            if isinstance(c, tuple):
                date_part, ot_part = c
                if isinstance(date_part, (pd.Timestamp, datetime.datetime)):
                    flat_cols.append(f"{ot_part} - {date_part.strftime('%d %b')}")
                else:
                    flat_cols.append(str(date_part) if date_part else str(ot_part))
            else:
                flat_cols.append(c)
        pivot_by_order.columns = flat_cols

        st.dataframe(pivot_by_order, use_container_width=True,hide_index=True)

    # ==============================================================================
    # F. تحليل الرصيد والتغطية
    # ==============================================================================
    st.markdown("---")
    st.subheader("📊 تحليل حرجية الرصيد ونسبة التغطية")

    if not result_df.empty:
        component_analysis = (
            merged_df
            .groupby([
                col("component"), col("component_desc"), col("component_uom"),
                col("current_stock"), col("component_order_type"),
                "BOM Level", col("mrp_controller"),
            ], as_index=False)
            .agg(
                Required_Qty=("Required Component Quantity", "sum"),
                Order_Types=("Order Type", lambda x: ", ".join(sorted(set(str(v) for v in x if pd.notna(v)))))
            )
            .rename(columns={
                "Required_Qty": "Required Component Quantity",
                "Order_Types": "Order Type",
            })
        )


        # 🔹 تنظيف وتحويل الأعمدة الرقمية
        numeric_cols = [col("current_stock"), "Required Component Quantity"]

        for c in numeric_cols:
                component_analysis[c] = component_analysis[c].astype(str).str.strip()
                component_analysis[c] = component_analysis[c].str.replace(r'[^\d\.]', '', regex=True)
                component_analysis[c] = pd.to_numeric(component_analysis[c], errors='coerce')

        # 🔹 حساب نسبة التغطية + تحويل الناتج + التقريب
        component_analysis["Coverage Percentage"] = pd.to_numeric(
                component_analysis[col("current_stock")] /
                component_analysis["Required Component Quantity"].replace(0, pd.NA) * 100,
                errors='coerce'
        ).round(1).fillna(0)



        component_analysis["Coverage Status"] = component_analysis["Coverage Percentage"].apply(
            lambda x: "🟢 كافية" if x >= 100 else ("🟡 جزئية" if x >= 50 else "🔴 غير كافية")
        )
        component_analysis["Priority"] = component_analysis.apply(
            lambda row: "🔥 عاجل" if row["Coverage Percentage"] < 30 and row["Required Component Quantity"] > 1000
            else ("⚠️ متوسط" if row["Coverage Percentage"] < 50 else "✅ منخفض"),
            axis=1
        )

        # --- فلاتر ---
        col1, col2, col3 = st.columns(3)
        with col1:
            mrp_opts = sorted(component_analysis[col("mrp_controller")].dropna().unique())
            selected_mrp = st.multiselect("🔍 MRP Controller:", options=mrp_opts, default=mrp_opts)
        with col2:
            ot_opts = sorted(component_analysis[col("component_order_type")].dropna().unique())
            selected_ot = st.multiselect("🔍 نوع طلب المكون:", options=ot_opts, default=ot_opts)
        with col3:
            lv_opts = sorted(component_analysis["BOM Level"].dropna().unique())
            selected_lv = st.multiselect("🔍 المستوى الهرمي:", options=lv_opts, default=lv_opts)

        filtered_analysis = component_analysis[
            component_analysis[col("mrp_controller")].isin(selected_mrp) &
            component_analysis[col("component_order_type")].isin(selected_ot) &
            component_analysis["BOM Level"].isin(selected_lv)
        ]

        st.dataframe(filtered_analysis.sort_values("Coverage Percentage"), use_container_width=True,hide_index=True)

        # إحصائيات التغطية
        tc  = max(len(filtered_analysis), 1)
        sc  = len(filtered_analysis[filtered_analysis["Coverage Percentage"] >= 100])
        pc  = len(filtered_analysis[(filtered_analysis["Coverage Percentage"] >= 50) & (filtered_analysis["Coverage Percentage"] < 100)])
        ic  = len(filtered_analysis[filtered_analysis["Coverage Percentage"] < 50])
        crt = len(filtered_analysis[filtered_analysis["Priority"] == "🔥 عاجل"])

        st.markdown(f"""
        <div style="direction:rtl; text-align:right; font-size:18px;">
        <span style="font-size:20px; color:#1976d2;">📈 <b>إحصائيات نسبة التغطية</b></span><br><br>
        <ul style="list-style-type:none; padding-right:0;">
            <li>🟢 <b>{sc}</b> مكونات تغطية كافية ({sc/tc*100:.1f}%)</li>
            <li>🟡 <b>{pc}</b> مكونات تغطية جزئية ({pc/tc*100:.1f}%)</li>
            <li>🔴 <b>{ic}</b> مكونات تغطية غير كافية ({ic/tc*100:.1f}%)</li>
            <li>🔥 <b style="color:red;">{crt}</b> مكونات حرجة تحتاج اهتمام عاجل</li>
        </ul>
        </div>
        """, unsafe_allow_html=True)

        # رسوم بيانية
        fig_pie = px.pie(
            filtered_analysis,
            names="Coverage Status",
            title="توزيع المكونات حسب حالة التغطية",
            color="Coverage Status",
            color_discrete_map={"🟢 كافية": "green", "🟡 جزئية": "orange", "🔴 غير كافية": "red"}
        )
        st.plotly_chart(fig_pie, use_container_width=True)

        top_critical = filtered_analysis.nsmallest(10, "Coverage Percentage").copy()
        if not top_critical.empty:
            top_critical[col("component")]      = top_critical[col("component")].astype(str)
            top_critical[col("component_desc")] = top_critical[col("component_desc")].astype(str)
            top_critical["Short_Label"] = (
                top_critical[col("component")] + " - " +
                top_critical[col("component_desc")].str[:25]
            )
            top_critical = top_critical.sort_values("Required Component Quantity", ascending=True)

            fig_crit = px.bar(
                top_critical,
                y="Short_Label",
                x="Required Component Quantity",
                color="Coverage Percentage",
                orientation='h',
                title="أقل 10 مكونات في نسبة التغطية",
                labels={"Required Component Quantity": "كمية الطلب", "Short_Label": "المكون", "Coverage Percentage": "نسبة التغطية %"},
                color_continuous_scale="RdYlGn_r"
            )
            fig_crit.update_layout(height=450)
            st.plotly_chart(fig_crit, use_container_width=True)

        # تحليل حسب MRP Controller والمستوى
        if len(selected_mrp) > 1:
            fig_sunburst = px.sunburst(
                filtered_analysis,
                path=[col("mrp_controller"), "BOM Level", "Coverage Status"],
                values="Required Component Quantity",
                title="توزيع الاحتياج حسب MRP Controller والمستوى وحالة التغطية"
            )
            st.plotly_chart(fig_sunburst, use_container_width=True)

        # تحليل حسب نوع الطلب
        fig_ot = px.pie(
            filtered_analysis,
            names=col("component_order_type"),
            title="توزيع المكونات حسب نوع الطلب"
        )
        st.plotly_chart(fig_ot, use_container_width=True)

    # ==============================================================================
    # G. Component in BOMs — النمطي التراكمي لكل مكون داخل منتج تام = 1 وحدة
    # ==============================================================================
    st.markdown("---")
    st.subheader("📋 قائمة الموديلات التي تستخدم كل مكون (نمطي لكل منتج تام = 1)")
 
    if not result_df.empty:
        # نُنشئ plan_unit: صف واحد لكل (Material, material_desc, Order Type) بكمية = 1
        unit_plan = (
            plan_melted[[col("material"), col("material_desc"), col("order_type")]]
            .drop_duplicates()
            .copy()
        )
        unit_plan["Planned Quantity"] = 1
        unit_plan["Date"] = pd.Timestamp("2000-01-01")   # تاريخ وهمي ثابت
 
        # نُشغّل explosion بكمية = 1 → يعطي النمطي التراكمي لكل منتج
        unit_result = bom_explosion(unit_plan, component_df)
 
        if not unit_result.empty:
            # 🔹 المفتاح: Material + Order Type فقط (بدون material_desc)
            # السبب: material_desc في unit_result يأتي من bom_explosion وقد يكون فارغاً
            # مما يُفشل الدمج ويُعيد plan_qty = NaN → 0
            plan_qty_map = (
                plan_melted.groupby([col("material"), col("order_type")])["Planned Quantity"]
                .sum()
                .reset_index()
                .rename(columns={"Planned Quantity": "plan_qty"})
            )
            # نُحضّر material_desc الصحيح من plan_melted بشكل منفصل
            mat_desc_map = (
                plan_melted[[col("material"), col("material_desc")]]
                .drop_duplicates(subset=[col("material")])
                .copy()
            )

            # 🔹 توحيد الأنواع
            unit_result[col("material")] = unit_result[col("material")].astype(str)
            unit_result["Order Type"]    = unit_result["Order Type"].astype(str)
            plan_qty_map[col("material")]   = plan_qty_map[col("material")].astype(str)
            plan_qty_map[col("order_type")] = plan_qty_map[col("order_type")].astype(str)
            mat_desc_map[col("material")]   = mat_desc_map[col("material")].astype(str)

            # 🔹 دمج plan_qty بـ Material + Order Type ← يضمن إيجاد الكمية دائماً
            unit_result = unit_result.merge(
                plan_qty_map,
                left_on=[col("material"), "Order Type"],
                right_on=[col("material"), col("order_type")],
                how="left",
                suffixes=("", "_plan")
            )
            # 🔹 دمج material_desc الصحيح (يُستخدم في رأس العمود فقط)
            if col("material_desc") in unit_result.columns:
                unit_result.drop(columns=[col("material_desc")], inplace=True)
            unit_result = unit_result.merge(
                mat_desc_map,
                on=col("material"),
                how="left"
            )
 
            # رأس العمود: كود المنتج , الكمية الفعلية , وصفه (نوع الطلب)
            unit_result["model_info"] = (
                unit_result[col("material")].astype(str) + " , " +
                unit_result["plan_qty"].fillna(0).round(0).astype(int).astype(str) + " , " +
                unit_result[col("material_desc")].astype(str).fillna("") + " (" +
                unit_result["Order Type"].astype(str).fillna("") + ")"
            )
 
            # ✅ نُلغي BOM Level من الـ pivot — نجمع كل المستويات في صف واحد
            pivot_index = [col("component"), col("component_desc"),
                           col("mrp_controller"), col("component_uom")]
            pivot_index = [c for c in pivot_index if c in unit_result.columns]
 
            component_bom_pivot = unit_result.pivot_table(
                index=pivot_index,
                columns="model_info",
                values="Required Component Quantity",
                aggfunc="sum"
            ).reset_index()
            component_bom_pivot.columns.name = None
 
            st.dataframe(component_bom_pivot.round(3).fillna(""), use_container_width=True)
        else:
            component_bom_pivot = pd.DataFrame()
            st.info("لا توجد بيانات لعرضها في جدول النمطي.")
    else:
        component_bom_pivot = pd.DataFrame()
        st.info("لا توجد نتائج BOM لعرض النمطي.")
 
    # ==============================================================================
    # H. جدول الكميات الشهرية + الرسم البياني
    # نستخدم plan_df_orig لأن plan_df تم تحويل أعمدة التواريخ فيه إلى نصوص
    # ==============================================================================
    st.markdown("---")
    if date_cols:
        orders_summary = plan_df_orig.melt(
            id_vars=[col("material"), col("material_desc"), col("order_type")],
            value_vars=date_cols,
            var_name="Month",
            value_name="Quantity"
        )
        orders_summary["Quantity"] = pd.to_numeric(orders_summary["Quantity"], errors="coerce").fillna(0)
        try:
            orders_summary["Month"] = pd.to_datetime(orders_summary["Month"]).dt.month_name()
        except Exception:
            pass

        orders_grouped = (
            orders_summary
            .groupby(["Month", col("order_type")])
            .agg({"Quantity": "sum"})
            .reset_index()
        )
        pivot_monthly = orders_grouped.pivot_table(
            index="Month", columns=col("order_type"),
            values="Quantity", aggfunc="sum", fill_value=0
        ).reset_index()

        if "E" not in pivot_monthly.columns: pivot_monthly["E"] = 0
        if "L" not in pivot_monthly.columns: pivot_monthly["L"] = 0
        pivot_monthly["الإجمالي"] = pivot_monthly["E"] + pivot_monthly["L"]
        total_sum = pivot_monthly["الإجمالي"].sum()
        if total_sum > 0:
            pivot_monthly["E%"] = (pivot_monthly["E"] / pivot_monthly["الإجمالي"] * 100).round(1).astype(str) + "%"
            pivot_monthly["L%"] = (pivot_monthly["L"] / pivot_monthly["الإجمالي"] * 100).round(1).astype(str) + "%"
        else:
            pivot_monthly["E%"] = pivot_monthly["L%"] = "0.0%"

        month_order = {m: i for i, m in enumerate(calendar.month_name) if m}
        pivot_monthly = pivot_monthly.sort_values(
            by="Month", key=lambda x: x.map(lambda v: month_order.get(v, 99))
        )

        st.subheader("📊 توزيع الكميات الشهرية حسب نوع الأمر")
        html_table = (
            "<table border='1' style='border-collapse:collapse;width:100%;text-align:center;'>"
            "<tr style='background-color:#1976d2;color:white;'>"
            "<th>الشهر</th><th>E</th><th>L</th><th>الإجمالي</th><th>E%</th><th>L%</th></tr>"
        )
        for _, row in pivot_monthly.iterrows():
            html_table += (
                f"<tr><td style='color:blue;font-weight:bold;'>{row['Month']}</td>"
                f"<td>{int(row.get('E',0)):,}</td><td>{int(row.get('L',0)):,}</td>"
                f"<td>{int(row.get('الإجمالي',0)):,}</td>"
                f"<td>{row.get('E%','')}</td><td>{row.get('L%','')}</td></tr>"
            )
        html_table += "</table>"
        st.markdown(f"<div style='direction:rtl;'>{html_table}</div>", unsafe_allow_html=True)

        fig_bar = px.bar(
            pivot_monthly, x="Month", y=["E", "L"],
            barmode="group", text_auto=True,
            title="رسم بياني لتوزيع الكميات الشهرية",
            labels={"value": "الكمية", "variable": "نوع الأمر", "Month": "الشهر"},
            template="streamlit"
        )
        st.plotly_chart(fig_bar, use_container_width=True)

    # ==============================================================================
    # I. إعداد ملف الـ Summary للتصدير
    # ==============================================================================
    coverage_stats_export = []
    if not result_df.empty:
        tc2 = max(len(component_analysis), 1)
        sc2  = len(component_analysis[component_analysis["Coverage Percentage"] >= 100])
        pc2  = len(component_analysis[(component_analysis["Coverage Percentage"] >= 50) & (component_analysis["Coverage Percentage"] < 100)])
        ic2  = len(component_analysis[component_analysis["Coverage Percentage"] < 50])
        crt2 = len(component_analysis[component_analysis["Priority"] == "🔥 عاجل"])
        coverage_stats_export = [
            ["🟢 مكونات تغطية كافية", sc2, f"{sc2/tc2*100:.1f}%"],
            ["🟡 مكونات تغطية جزئية", pc2, f"{pc2/tc2*100:.1f}%"],
            ["🔴 مكونات تغطية غير كافية", ic2, f"{ic2/tc2*100:.1f}%"],
            ["🔥 مكونات حرجة", crt2, ""],
        ]

    # ── بيانات الكميات الشهرية للـ Summary ──────────────────────────────────
    monthly_summary_rows = []
    if date_cols:
        monthly_summary_rows = [["", "", ""], ["📅 الكميات الشهرية", "", ""]]
        for _, mrow in pivot_monthly.iterrows():
            monthly_summary_rows.append([
                mrow["Month"],
                int(mrow.get("الإجمالي", 0)),
                f"E: {int(mrow.get('E',0)):,}  |  L: {int(mrow.get('L',0)):,}"
            ])
        e_total = int(pivot_monthly.get("E", pd.Series([0])).sum())
        l_total = int(pivot_monthly.get("L", pd.Series([0])).sum())
        grand   = e_total + l_total
        monthly_summary_rows.append(["الإجمالي الكلي", grand, f"E: {e_total:,}  |  L: {l_total:,}"])

    summary_data = [
        ["📌 ملخص نتائج الخطة", "", ""],
        ["موديلات بالخطة", total_models, ""],
        ["مكونات فريدة", total_components, ""],
        ["سطور BOM", total_boms, ""],
        ["مكونات بدون MRP Controller", empty_mrp_count, ""],
        ["مكونات بأكثر من وحدة", total_diff_uom, diff_uom_str],
        ["منتجات بالخطة بدون BOM", total_missing_boms, ""],
        ["", "", ""],
        ["مكونات شراء (F)", purchase_count, ""],
        ["مكونات تصنيع (E)", manufacturing_count, ""],
        ["مكونات غير محددة", undefined_count, ""],
        ["", "", ""],
        ["📈 إحصائيات التغطية", "", ""],
        *coverage_stats_export,
        *monthly_summary_rows,
        ["", "", ""],
        ["تاريخ الإنشاء", datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), ""],
    ]
    summary_df = pd.DataFrame(summary_data, columns=["البند", "القيمة", "ملاحظات"])

    # تنسيق plan_df للتصدير
    plan_df_export = plan_df.copy()
    plan_df_export.columns = [
        c.strftime("%d %b") if isinstance(c, (datetime.datetime, pd.Timestamp)) else c
        for c in plan_df_export.columns
    ]
    # ==============================================================================
    # J. تصدير Excel — مع اختيار المستخدم للأوراق ولـ MRP Controller
    # ==============================================================================
    st.markdown("---")
    st.subheader("📤 تصدير النتائج إلى Excel")

    # ── 1. اختيار MRP Controller (يؤثر على كل الأوراق التي تحتوي العمود) ──
    mrp_controller_col = col("mrp_controller")   # "MRP Controller"
    if not mrp_df.empty and mrp_controller_col in mrp_df.columns:
        mrp_options = sorted(mrp_df[mrp_controller_col].dropna().unique().tolist())
    elif not result_df.empty and mrp_controller_col in result_df.columns:
        mrp_options = sorted(result_df[mrp_controller_col].dropna().unique().tolist())
    else:
        mrp_options = []

    if mrp_options:
        # عنوان كبير وأزرق وبولد
        st.markdown(
            '<p style="font-size:18px; color:blue; font-weight:bold;">👤 اختر MRP Controller المراد تصديرهم (يُطبَّق على جميع الأوراق التي تحتوي العمود):</p>',
            unsafe_allow_html=True
        )

        # Multiselect بدون Label لأنه موجود في الـ HTML أعلاه
        selected_mrp = st.multiselect(
            "",
            options=mrp_options,
            default=mrp_options
        )
    else:
        selected_mrp = []

    # ── 2. تعريف الأوراق المتاحة ──────────────────────────────────────────
    available_sheets = {
        "📋 الخطة الأصلية (Original_Plan)":        ("Original_Plan",           True),
        "📌 الملخص (Summary)":                     ("Summary",                 True),
        "📅 الاحتياج بالتاريخ (Need_By_Date)":      ("Need_By_Date",            not result_df.empty),
        "📦 الاحتياج بنوع الأمر (Need_By_Order)":   ("Need_By_Order_Type",      not result_df.empty),
        "🔍 تحليل التغطية (Stock_Coverage)":        ("Stock_Coverage_Analysis", not result_df.empty),
        "🌳 BOM الكامل (BOM_All_Levels)":           ("BOM_All_Levels",          not result_df.empty),
        "📊 النمطي لكل منتج (Component_in_BOMs)":   ("Component_in_BOMs",       not component_bom_pivot.empty),
        "🗂️ المكونات الأصلية (Original_Component)": ("Original_Component",      True),
    }

    # أوراق مفعّلة افتراضيًا
    default_checked = {"Original_Plan", "Need_By_Date", "Component_in_BOMs"}

    # ── 3. Checkboxes في عمودين ───────────────────────────────────────────
    st.markdown(
        '<h1 style="font-size:32px; color:blue; font-weight:bold;">اختر الأوراق التي تريد تصديرها:</h1>',
        unsafe_allow_html=True
    )

    col1, col2 = st.columns(2)
    selected_sheets = {}
    sheet_items = list(available_sheets.items())
    for i, (label, (sheet_name, available)) in enumerate(sheet_items):
        target_col = col1 if i % 2 == 0 else col2
        with target_col:
            if available:
                selected_sheets[sheet_name] = st.checkbox(
                    label,
                    value=(sheet_name in default_checked),
                    key=f"sheet_{sheet_name}"
                )
            else:
                st.checkbox(
                    label + " *(غير متاح)*",
                    value=False,
                    disabled=True,
                    key=f"sheet_{sheet_name}_dis"
                )





    # ── 4. زر التصدير ────────────────────────────────────────────────────
    if st.button("🗜️ اضغط هنا لإنشاء النسخة الكاملة"):
        chosen = [k for k, v in selected_sheets.items() if v]
        if not chosen:
            st.warning("⚠️ لم تختر أي ورقة للتصدير.")
        else:
            with st.spinner("⏳ جاري إنشاء ملف Excel..."):
                current_date = datetime.datetime.now().strftime("%d_%b_%Y")
                excel_buffer = BytesIO()

                # ── خريطة الأوراق ───────────────────────────────────────
                sheet_data_map = {
                    "Original_Plan":           plan_df_export,
                    "Summary":                 summary_df,
                    "Need_By_Date":            pivot_by_date          if not result_df.empty        else pd.DataFrame(),
                    "Need_By_Order_Type":      pivot_by_order         if not result_df.empty        else pd.DataFrame(),
                    "Stock_Coverage_Analysis": component_analysis     if not result_df.empty        else pd.DataFrame(),
                    "BOM_All_Levels":          merged_df              if not result_df.empty        else pd.DataFrame(),
                    "Component_in_BOMs":       component_bom_pivot    if not component_bom_pivot.empty else pd.DataFrame(),
                    "Original_Component":      component_df_orig,
                    "MRP_Controller":          mrp_df                 if not mrp_df.empty           else pd.DataFrame(),
                }

                # ── تطبيق فلتر MRP Controller على كل ورقة تحتوي العمود ─
                if selected_mrp:
                    for sname, sdf in sheet_data_map.items():
                        if not sdf.empty and mrp_controller_col in sdf.columns:
                            sheet_data_map[sname] = sdf[sdf[mrp_controller_col].isin(selected_mrp)]

                # ── الكتابة ──────────────────────────────────────────────
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    for sheet_name in chosen:
                        df_to_write = sheet_data_map.get(sheet_name, pd.DataFrame())
                        if not df_to_write.empty:
                            df_to_write.to_excel(writer, sheet_name=sheet_name, index=False)

                excel_buffer.seek(0)
                st.download_button(
                    label="📊 تحميل ملف Excel الكامل",
                    data=excel_buffer,
                    file_name=f"MRP_Results_{current_date}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.balloons()
                st.success(f"✅ تم إنشاء الملف بنجاح — {len(chosen)} ورقة: {', '.join(chosen)}")

# --- التذييل ---
st.markdown("""
<hr>
<div style="text-align:center; direction:rtl; font-size:14px; color:gray;">
    ✨ تم التنفيذ بواسطة <b>م / رضا رشدي</b> — جميع الحقوق محفوظة © 2026 ✨
</div>
""", unsafe_allow_html=True)
