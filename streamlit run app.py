import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="🔥 MRP BOM Explosion  ", layout="wide")
st.title("🔥 MRP Tool -  MRP BOM Explosion استخراج الاحتياجات من المكونات لخطة أنتاج ")

# رفع ملف Excel
uploaded_file = st.file_uploader("اختر ملف Excel يحتوي على أوراق plan و Component", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    if "plan" not in xls.sheet_names or "Component" not in xls.sheet_names:
        st.error("❌ الملف لا يحتوي على أوراق plan و Component")
        st.stop()

    # قراءة أوراق Excel
    plan_df = pd.read_excel(xls, sheet_name="plan")
    component_df = pd.read_excel(xls, sheet_name="Component")

    # -----------------------------------------------
    # 1. إزالة الصفوف المكررة وحفظ الأصلية
    original_plan_df = plan_df.copy() # نسخة من الخطة الأصلية للحفظ
    component_df_orig = component_df.copy() # نسخة من بيانات المكونات الأصلية (لتقرير Component in BOMs)

    plan_df = plan_df.drop_duplicates()
    component_df = component_df.drop_duplicates()

    # 2. توحيد الوحدات للوزن (جرام → كيلوجرام)
    def normalize_units(df):
        df = df.copy()
        uom_col = "Component UoM"
        qty_col = "Component Quantity"
        grams_units = ["g", "gram", "grams", "gm", "جرام", "غ"]
        mask = df[uom_col].astype(str).str.lower().isin(grams_units)
        df.loc[mask, qty_col] = df.loc[mask, qty_col] / 1000
        df.loc[mask, uom_col] = "kg"
        return df

    component_df = normalize_units(component_df) 

    # 3. تحويل خطة الإنتاج من wide إلى long
    date_cols = [c for c in plan_df.columns if c not in ["Material", "Material Description", "Order Type"]]
    plan_melted = plan_df.melt(
        id_vars=["Material", "Material Description", "Order Type"],
        value_vars=date_cols,
        var_name="Date",
        value_name="Planned Quantity"
    )
    plan_melted["Date"] = pd.to_datetime(plan_melted["Date"], errors='coerce')
    plan_melted["Planned Quantity"] = pd.to_numeric(plan_melted["Planned Quantity"], errors='coerce').fillna(0)
    plan_melted = plan_melted[plan_melted["Planned Quantity"] > 0]


    # -----------------------------------------------
    # استخراج قائمة المنتجات النهائية من الخطة
    top_level_materials = plan_df["Material"].unique()
    # -----------------------------------------------

    
    # ===============================================
    # 1. دالة التفجير الهرمي حسب التاريخ فقط
    # ===============================================
    def explode_bom(parent_material, qty, date, comp_df, results):
        children = comp_df[comp_df["Parent Material"] == parent_material]
        if children.empty:
            return
        for _, row in children.iterrows():
            child = row["Component"]
            per_unit = row["Component Quantity"]
            required_qty = qty * per_unit
            results.append({
                "Component": child,
                "Component Description": row.get("Component Description", ""),
                "UoM": row.get("Component UoM", ""),
                "MRP Contor": row.get("MRP Controller", ""),
                "Date": date,
                "Required Qty": required_qty
            })
            explode_bom(child, required_qty, date, comp_df, results)

    # تنفيذ التفجير للتجميع حسب التاريخ فقط (التقرير الأول)
    results_date_only = []
    for _, plan_row in plan_melted.iterrows():
        explode_bom(plan_row["Material"], plan_row["Planned Quantity"], plan_row["Date"], component_df, results_date_only)

    final_df = pd.DataFrame(results_date_only)

    # ===============================================
    # 2. دالة التفجير الهرمي حسب الشهر ونوع الطلب
    # ===============================================
    def explode_recursive_order(parent_material, qty, order_type, order_date, comp_df, results, path):
        children = comp_df[comp_df["Parent Material"] == parent_material] 
        if children.empty:
            return
        for _, crow in children.iterrows():
            child_code = crow["Component"]
            if child_code in path:
                st.warning(f"❌ تم تجاهل المكون: {child_code} لتجنب حلقة تكرارية في BOM.")
                continue
                
            per_unit = crow.get("Component Quantity", 0.0)
            child_qty = qty * per_unit
            mrp_contor = crow.get("MRP Controller", "N/A") 
            
            results.append({
                "Component": child_code,
                "Component Description": crow.get("Component Description", ""),
                "Component UoM": crow.get("Component UoM", ""),
                "MRP Contor": mrp_contor, 
                "Order Type": order_type,
                "Month": pd.to_datetime(order_date).strftime("%b"), 
                "Required Component Quantity": child_qty
            })
            explode_recursive_order(child_code, child_qty, order_type, order_date, comp_df, results, path + [child_code])

    
    # تنفيذ التفجير لجميع المواد (التقرير الثاني)
    order_results = []
    for _, prow in plan_melted.iterrows():
        top_material = prow["Material"]
        plan_qty = prow["Planned Quantity"]
        order_type = prow.get("Order Type", "N/A")
        order_date = prow.get("Date", None)
        
        if plan_qty == 0 or pd.isna(order_date):
            continue
            
        explode_recursive_order(top_material, plan_qty, order_type, order_date, component_df, order_results, path=[top_material])

    order_df = pd.DataFrame(order_results)


    # ==============================================================================
    # 3. حساب جدول (Top-Level BOM) - خريطة المكونات في المنتج النهائي
    # ==============================================================================
 #   st.subheader("📋 التقرير الثالث: خريطة المكونات في المنتج النهائي (Top-Level BOM)")

    # 1. دمج بيانات الخطة (Plan) مع بيانات المكونات الأصلية (Component)
    plan_summary = plan_melted.groupby(
        ["Material", "Order Type"]
    )["Planned Quantity"].sum().reset_index()
    plan_summary.rename(columns={"Planned Quantity": "plan_qty"}, inplace=True)

    component_bom_merged = pd.merge(
        component_df_orig, 
        plan_summary, 
        left_on="Parent Material", 
        right_on="Material", 
        how="left"
    ).fillna({"plan_qty": 0, "Order Type": 'N/A'})

    # 🛑 التصفية: تشمل فقط الـ BOMs التي يكون فيها الأب هو أحد المنتجات النهائية في الخطة
    component_bom_filtered = component_bom_merged[
        component_bom_merged["Parent Material"].isin(top_level_materials)
    ].copy()


    # 2. إنشاء عمود تجميعي (model_info) للمحور الأفقي
    if not component_bom_filtered.empty:
        component_bom_filtered["model_info"] = (
            component_bom_filtered["Parent Material"].astype(str)
            + " ("
            + component_bom_filtered["plan_qty"].astype(int).astype(str)
            + " "
            + component_bom_filtered["Order Type"].astype(str)
            + ")"
        )

        # 3. إنشاء جدول محوري
        component_bom_pivot = component_bom_filtered.pivot_table(
            index=[
                "Component",
                "Component Description",
                "MRP Controller", 
                "Component UoM"
            ],
            columns="model_info",
            values="Component Quantity", 
            aggfunc="first",
            fill_value=0
        ).reset_index()

        component_bom_pivot.rename(columns={"MRP Controller": "MRP Contor"}, inplace=True)
        
        # 4. عرض الجدول في الواجهة
  #      st.dataframe(component_bom_pivot, use_container_width=True)
   #     st.markdown("---")
  #  else:
   #     component_bom_pivot = pd.DataFrame()
    #    st.warning("⚠️ لا توجد بيانات لتقرير خريطة المكونات في المنتج النهائي.")
        

    # -----------------------------------------------
    # عرض النتائج في الواجهة
    # -----------------------------------------------

    # التقرير الأول: Need By Date
    if not final_df.empty:
        agg_df = final_df.groupby(
            ["Component", "Component Description", "UoM", "MRP Contor", "Date"], as_index=False
        )["Required Qty"].sum()
        pivot_df_date = agg_df.pivot_table(
            index=["Component", "Component Description", "UoM", "MRP Contor"],
            columns="Date",
            values="Required Qty",
            fill_value=0
        ).reset_index()
        pivot_df_date.columns = [col.strftime("%d %b") if isinstance(col, pd.Timestamp) else col for col in pivot_df_date.columns]
 #       st.subheader("📆 التقرير الأول: احتياجات كل مكون حسب التاريخ (Need_By_Date)")
  #      st.dataframe(pivot_df_date)
    else:
        pivot_df_date = pd.DataFrame()


    # التقرير الثاني: Need By Order Type and Month
    if not order_df.empty:
 #       st.subheader("📊 التقرير الثاني: احتياجات كل مكون حسب نوع الطلب والشهر (Need_By_OrderType_Month)")
        result_order = order_df.groupby(
            ["Component", "Component Description", "Component UoM", "MRP Contor", "Order Type", "Month"],
            as_index=False
        )["Required Component Quantity"].sum()
        result_order["Order_Month"] = result_order["Month"] + " (" + result_order["Order Type"] + ")"
        pivot_by_order = result_order.pivot_table(
            index=["Component", "Component Description", "Component UoM", "MRP Contor"],
            columns="Order_Month",
            values="Required Component Quantity",
            aggfunc="sum",
            fill_value=0
        ).reset_index()
#        st.dataframe(pivot_by_order)
    else:
        pivot_by_order = pd.DataFrame()


    # ===============================================
    # تصدير النتائج Excel - مع جميع التقارير
    # ===============================================
    if not pivot_df_date.empty or not pivot_by_order.empty or not component_bom_pivot.empty:
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            
            # 1. التقرير الأول: احتياجات حسب التاريخ
            if not pivot_df_date.empty:
                pivot_df_date.to_excel(writer, sheet_name="Need_By_Date", index=False)
            
            # 2. التقرير الثاني: احتياجات حسب نوع الطلب والشهر
 #           if not pivot_by_order.empty:
  #              pivot_by_order.to_excel(writer, sheet_name="Need_By_OrderType", index=False)
                
            # 3. التقرير الثالث: خريطة المكونات في المنتج النهائي
 #           if not component_bom_pivot.empty:
  #              component_bom_pivot.to_excel(writer, sheet_name="Top_Level_BOM", index=False) 
                
            # 4. الخطة الأصلية
            original_plan_df.to_excel(writer, sheet_name="Original_Plan", index=False) 
            
        buffer.seek(0)

        st.download_button(
            label="📥 تحميل جميع النتائج Excel",
            data=buffer,
            file_name="MRP_Explosion_Reports.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
