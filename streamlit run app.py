import streamlit as st
import pandas as pd
import datetime
from io import BytesIO
import calendar
import plotly.express as px
import sys # ✅ إضافة مكتبة النظام
# ==========================================================
# دالة حساب الـ MRP متعدد المستويات (Multi-Level MRP)
# ✅ تم تعديل هذه الدالة لتطبيق أولوية الوحدة "القطعة"
# ==========================================================
def calculate_multi_level_mrp(plan_df, component_df):
    
    # 1. تجهيز الـ BOMs لتحديد المكونات المصنعة داخلياً (التي تحتاج MRP)
    manufactured_components = set(component_df["Material"].unique())
    
    # 2. تجهيز الخطة الأولية كطلب (Initial Demand)
    date_cols = [c for c in plan_df.columns if isinstance(c, (datetime.datetime, pd.Timestamp))]
    
    # دمج بيانات الخطة (Plan) وتحويلها من أعمدة إلى صفوف (melt)
    demand_df = plan_df.melt(
        id_vars=["Material", "Material Description", "Order Type"],
        var_name="Date",
        value_name="Required Quantity"
    )
    demand_df["Date"] = pd.to_datetime(demand_df["Date"], errors='coerce')
    demand_df = demand_df.dropna(subset=["Date", "Required Quantity"])
    demand_df = demand_df[demand_df["Required Quantity"] > 0] # تجاهل الكميات الصفرية

    final_requirements = pd.DataFrame()
    
    # المتغير الذي يحمل طلبات المستوى الحالي (يبدأ بخطة المنتج النهائي)
    current_demand = demand_df.rename(columns={'Material': 'Parent'})
    
    # بدء عملية التكرار (Roll-Down)
    while not current_demand.empty:
        
        merged = pd.merge(
            current_demand, 
            component_df, 
            left_on='Parent', 
            right_on='Material', 
            how='inner'
        )

        merged['Calculated Quantity'] = merged['Required Quantity'] * merged['Component Quantity']
        
        # *******************************************************************
        # ✅ تطبيق أولوية "القطعة" قبل التجميع في كل مستوى
        # *******************************************************************
        
        # 1. ترتيب البيانات: إعطاء الأولوية للوحدة 'القطعة' (القيمة 0)
        merged.loc[:, 'Priority_Order'] = merged['Component UoM'].apply(lambda x: 0 if x == 'القطعة' else 1)
        # *******************************************************************        
        # 2. تجميع متطلبات المكونات الأبناء
        # *******************************************************************

        requirements_for_level = merged.sort_values(
            ['Date', 'Component', 'Priority_Order']
        ).groupby(
            ["Date", "Component", "Component UoM"] # ✅ التجميع بناءً على المكون والتاريخ والوحدة
        ).agg(
            {
                'Calculated Quantity': 'sum', # جمع الكميات المطلوبة
                'Component Description': 'first',
            }
        ).reset_index()        
        requirements_for_level = requirements_for_level.rename(
            columns={'Calculated Quantity': 'Required Component Quantity'}
        )
        
        # حذف عمود الأولوية المؤقت
        requirements_for_level = requirements_for_level.drop(columns=['Priority_Order'], errors='ignore')

        # *******************************************************************
        
        # إضافة متطلبات هذا المستوى إلى جدول المتطلبات النهائية
        final_requirements = pd.concat([final_requirements, requirements_for_level])

        # تجهيز الطلب (Demand) للمستوى التالي:
        next_level_demand_components = requirements_for_level[
            requirements_for_level['Component'].isin(manufactured_components)
        ]
        
        current_demand = next_level_demand_components.rename(
            columns={'Component': 'Parent', 'Required Component Quantity': 'Required Quantity'}
        )
        
        # حذف الأعمدة غير المطلوبة للحساب التالي
        current_demand = current_demand.drop(columns=['Component Description', 'Component UoM'], errors='ignore')
        
        if current_demand.empty:
            break
            
    # التجميع النهائي: جمع كل متطلبات المكونات
    final_mrp_result = final_requirements.groupby(['Component', 'Component Description', 'Component UoM', 'Date'])['Required Component Quantity'].sum().reset_index()
    return final_mrp_result

# إعداد الصفحة
st.set_page_config(page_title="🔥 MRP Tool", page_icon="📂", layout="wide")
st.subheader("📂 برنامج أستخراج وحفظ نتائج الـ MRP Need_By_Date Multi level")
st.markdown(
    "<p style='font-size:1.0em; font-weight:bold;'>💡 اختر ملف الخطة الشهرية Excel</p>",
    unsafe_allow_html=True
)

uploaded_file = st.file_uploader("", type=["xlsx"])

if uploaded_file:
    with st.spinner("⏳ جاري معالجة البيانات ----- انتظر قليلا.....⏳"):
        
        # -------------------------------
        # قراءة شيتات Excel
        # -------------------------------
        xls = pd.ExcelFile(uploaded_file, engine='openpyxl')
        
        # 1. التحقق من وجود جميع الأوراق المطلوبة
        required_sheets = ["plan", "Component"]
        missing_sheets = [sheet for sheet in required_sheets if sheet not in xls.sheet_names]
        
        if missing_sheets:
            st.error(f"❌ الملف لا يحتوي على الأوراق المطلوبة: {', '.join(missing_sheets)}")
            st.stop()
            
        plan_df = xls.parse("plan")
        component_df = xls.parse("Component")
        mrp_df = xls.parse("MRP Contor") if "MRP Contor" in xls.sheet_names else pd.DataFrame()

        # 2. التحقق من صحة البيانات الأساسية
        if plan_df.empty:
            st.error("❌ جدول الخطة فارغ. يرجى التحقق من الملف.")
            st.stop()

        if component_df.empty:
            st.error("❌ جدول المكونات فارغ. يرجى التحقق من الملف.")
            st.stop()

        # 3. التحقق من الأعمدة الأساسية في جدول الخطة:
        required_plan_columns = ["Material", "Material Description", "Order Type"]
        missing_plan_columns = [col for col in required_plan_columns if col not in plan_df.columns]
        if missing_plan_columns:
            st.error(f"❌ جدول الخطة لا يحتوي على الأعمدة المطلوبة: {', '.join(missing_plan_columns)}")
            st.stop()

        # 4. التحقق من الأعمدة الأساسية في جدول المكونات:
        # ✅ تم إضافة "Component UoM" كعمود مطلوب
        required_component_columns = ["Material", "Component", "Component Quantity", "Component UoM"]
        missing_component_columns = [col for col in required_component_columns if col not in component_df.columns]
        if missing_component_columns:
            st.error(f"❌ جدول المكونات لا يحتوي على الأعمدة المطلوبة: {', '.join(missing_component_columns)}. يجب إضافة عمود الوحدة (Component UoM).")
            st.stop()

        # *******************************************************************
        # تجهيز البيانات الأساسية
        # *******************************************************************
        plan_melted = plan_df.melt(
            id_vars=["Material", "Material Description", "Order Type"],
            var_name="Date",
            value_name="Planned Quantity"
        )
        plan_melted["Date"] = pd.to_datetime(plan_melted["Date"], errors='coerce')
        merged_df = pd.merge(plan_melted, component_df, on="Material", how="left")
        merged_df["Required Component Quantity"] = merged_df["Planned Quantity"] * merged_df["Component Quantity"]
        
        # إزالة الصفوف ذات الكمية المخططة الصفرية
        merged_df = merged_df[merged_df["Planned Quantity"] > 0]
        
        # *******************************************************************
        # حساب Multi-Level MRP
        # *******************************************************************
        result_date_multi = calculate_multi_level_mrp(plan_df, component_df)

        # -------------------------------
        # ✅ تجهيز شيت "MRP_Result"
        # -------------------------------
        mrp_result_sheet = result_date_multi.copy()
        mrp_result_sheet['Date'] = pd.to_datetime(mrp_result_sheet['Date'], errors='coerce').dt.strftime("%d %b")

        mrp_result_pivot = mrp_result_sheet.pivot_table(
            index=["Component", "Component Description", "Component UoM"],
            columns="Date",
            values="Required Component Quantity",
            aggfunc="sum",
            fill_value=0
        ).reset_index()

        date_cols = [c for c in mrp_result_pivot.columns if c not in ["Component", "Component Description", "Component UoM"]]
        mrp_result_pivot["Total Qty"] = mrp_result_pivot[date_cols].sum(axis=1)

        ordered_cols = ["Component", "Component Description", "Total Qty", "Component UoM"]
        other_cols = [c for c in mrp_result_pivot.columns if c not in ordered_cols]
        mrp_result_pivot = mrp_result_pivot[ordered_cols + other_cols]

        if not mrp_df.empty and "Component" in mrp_df.columns and "MRP Contor" in mrp_df.columns:
            mrp_result_pivot = pd.merge(
                mrp_result_pivot,
                mrp_df[["Component", "MRP Contor"]],
                on="Component",
                how="left"
            )
        else:
            mrp_result_pivot["MRP Contor"] = "N/A"

        cols = mrp_result_pivot.columns.tolist()
        fixed_order = ["Component", "Component Description", "Total Qty", "Component UoM", "MRP Contor"]
        date_cols = [c for c in cols if c not in fixed_order]
        mrp_result_pivot = mrp_result_pivot[fixed_order + date_cols]

        
        # *******************************************************************
        # دمج نتائج الـ MRP مع جدول MRP Contor
        # *******************************************************************
        if not mrp_df.empty and "MRP Contor" in mrp_df.columns and "Component" in mrp_df.columns:
            mrp_contor_cols = mrp_df[["Component", "MRP Contor"]].drop_duplicates()
            
            result_date_multi = pd.merge(
                result_date_multi, 
                mrp_contor_cols, 
                on="Component", 
                how="left"
            )
        else:
            st.warning("⚠️ لم يتم العثور على ورقة 'MRP Contor' أو الأعمدة المطلوبة بها، لن يتم إضافة عمود 'MRP Contor'.")
            result_date_multi["MRP Contor"] = "N/A" 

        # تحويل عمود التاريخ إلى صيغة نصية (dd mmm)
        result_date_multi['Date'] = result_date_multi['Date'].dt.strftime("%d %b")
        
        # إنشاء الجدول المحوري
        pivot_by_date_multi = result_date_multi.pivot(
            index=["Component", "Component Description", "Component UoM", "MRP Contor"],
            columns="Date",
            values="Required Component Quantity"
        ).reset_index()

        # *******************************************************************
        # الملخص السريع (عرض فقط)
        # *******************************************************************
        total_models = plan_df["Material"].nunique()
        total_components = component_df["Component"].nunique()
        total_boms = len(component_df)
        empty_mrp_count = mrp_df["Component"].isna().sum() if not mrp_df.empty else 0

        diff_uom = component_df.groupby("Component")["Component UoM"].nunique()
        diff_uom = diff_uom[diff_uom > 1]
        total_diff_uom = len(diff_uom)

        if total_diff_uom > 0:
            diff_uom_str = ", ".join(map(str, diff_uom.index))
            diff_uom_color = "red"
        else:
            diff_uom_str = "لا يوجد"
            diff_uom_color = "green"

        missing_boms = set(plan_df["Material"]) - set(component_df["Material"])
        total_missing_boms = len(missing_boms)
        missing_boms_html = (
            f"<span style='color:red;'>{', '.join(map(str, missing_boms))}</span>"
            if missing_boms else "<span style='color:green;'>لا يوجد</span>"
        )

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

        # *******************************************************************
        # Need_By_Date
        # *******************************************************************
        result_date = merged_df.groupby(
            ["Component", "Component Description", "Component UoM", "Date"]
        )["Required Component Quantity"].sum().reset_index()

        pivot_by_date = result_date.pivot(
            index=["Component", "Component Description", "Component UoM"],
            columns="Date",
            values="Required Component Quantity"
        ).reset_index()

        if not mrp_df.empty:
            pivot_by_date = pd.merge(
                pivot_by_date,
                mrp_df[["Component", "MRP Contor"]],
                on="Component",
                how="left"
            )

            # إعادة ترتيب الأعمدة بحيث MRP Contor يكون العمود الثالث
            cols = pivot_by_date.columns.tolist()
            fixed_order = ["Component", "Component Description", "MRP Contor", "Component UoM"]
            other_cols = [c for c in cols if c not in fixed_order]
            pivot_by_date = pivot_by_date[fixed_order + other_cols]

        # تنسيق أسماء الأعمدة (التواريخ تبقى dd mmm)
        pivot_by_date.columns = [
            col.strftime("%d %b") if isinstance(col, pd.Timestamp) else col
            for col in pivot_by_date.columns
        ]

        # *******************************************************************
        # Need_By_Order Type
        # *******************************************************************
        result_order = merged_df.groupby(
            ["Component", "Component Description", "Component UoM", "Order Type", "Date"]
        )["Required Component Quantity"].sum().reset_index()

        pivot_by_order = result_order.pivot_table(
            index=["Component", "Component Description", "Component UoM"],
            columns=["Date", "Order Type"],
            values="Required Component Quantity",
            aggfunc="sum",
            fill_value=0
        ).reset_index()

        pivot_by_order.columns = [
            f"{col[1][0]} - {col[0].strftime('%d %b')}" if isinstance(col, tuple) and isinstance(col[0], pd.Timestamp)
            else col if isinstance(col, str) else col[0]
            for col in pivot_by_order.columns
        ]


        if not mrp_df.empty:
            merged_df = merged_df.merge(mrp_df[["Component", "MRP Contor"]], on="Component", how="left")

            component_bom_map = merged_df.groupby(
                ["MRP Contor", "Component", "Material"]
            ).agg({
                "Order Type": lambda x: ','.join(sorted(set(x))),
                "Planned Quantity": "sum"
            }).reset_index()

            component_bom_map["OrderType_Quantity"] = component_bom_map["Order Type"] + " (" + component_bom_map["Planned Quantity"].astype(str) + ")"

            component_bom_pivot = component_bom_map.pivot_table(
                index=["MRP Contor", "Component"],
                columns="Material",
                values="OrderType_Quantity",
                aggfunc=lambda x: ','.join(x),
                fill_value=""
            )


        # *******************************************************************
        # جدول الكميات الشهرية + الرسم البياني (لم يتم المساس بهذا الجزء)
        # *******************************************************************
        date_cols = [c for c in plan_df.columns if isinstance(c, (datetime.datetime, pd.Timestamp))]
        if date_cols:
            orders_summary = plan_df.melt(
                id_vars=["Material", "Order Type"],
                value_vars=date_cols,
                var_name="Month",
                value_name="Quantity"
            )
            orders_summary["Month"] = pd.to_datetime(orders_summary["Month"], errors="coerce")
            orders_summary = orders_summary.dropna(subset=["Month"])

            orders_grouped = orders_summary.groupby(
                [orders_summary["Month"].dt.month_name(), "Order Type"]
            ).agg({"Quantity": "sum"}).reset_index()

            pivot_df = orders_grouped.pivot_table(
                index="Month", columns="Order Type", values="Quantity", aggfunc="sum", fill_value=0
            ).reset_index()

            pivot_df["الإجمالي"] = pivot_df.sum(axis=1, numeric_only=True)
            pivot_df["E%"] = (pivot_df.get("E",0)/pivot_df["الإجمالي"]*100).round(1).astype(str) + "%"
            pivot_df["L%"] = (pivot_df.get("L",0)/pivot_df["الإجمالي"]*100).round(1).astype(str) + "%"

            month_order = {m:i for i,m in enumerate(calendar.month_name) if m}
            pivot_df["MonthOrder"] = pivot_df["Month"].map(month_order)
            pivot_df = pivot_df.sort_values("MonthOrder").drop(columns="MonthOrder")

                       # عرض HTML منسق RTL
            st.subheader("🧩 توزيع الكميات الشهرية حسب نوع الأمر👇")
            html_table = "<table border='1' style='border-collapse: collapse; width:100%; text-align:center; color:black;'>"
            html_table += "<tr style='background-color:#d9d9d9; color:blue;'><th>الشهر</th><th>E</th><th>L</th><th>الإجمالي</th><th>E%</th><th>L%</th></tr>"

            for idx, row in pivot_df.iterrows():
                bg_color = "#f2f2f2" if idx % 2 == 0 else "#ffffff"
                html_table += f"<tr style='background-color:{bg_color};'>"
                html_table += f"<td style='color:blue;'>{row['Month']}</td>"
                html_table += f"<td>{int(row.get('E',0))}</td>"
                html_table += f"<td>{int(row.get('L',0))}</td>"
                html_table += f"<td>{int(row.get('الإجمالي',0))}</td>"
                html_table += f"<td>{row.get('E%','')}</td>"
                html_table += f"<td>{row.get('L%','')}</td>"
                html_table += "</tr>"

            html_table += "</table>"
            st.markdown(f"<div style='direction:rtl;'>{html_table}</div>", unsafe_allow_html=True)


            st.subheader("👇 رسم بياني للكميات 👇")

            # استبعاد عمود "الإجمالي" لو موجود
            numeric_cols = [c for c in pivot_df.columns if c not in ["Month", "الإجمالي"]]

            # اختيار الأعمدة الرقمية فقط
            numeric_cols = [c for c in numeric_cols if pd.api.types.is_numeric_dtype(pivot_df[c])]

            # تحويل القيم لأرقام (في حالة وجود نصوص تتحول NaN)
            pivot_df[numeric_cols] = pivot_df[numeric_cols].apply(pd.to_numeric, errors="coerce")

            # رسم العمودى
            fig = px.bar(
                pivot_df,
                x="Month",
                y=numeric_cols,
                barmode="group",
                text_auto=True,
                title="توزيع الكميات حسب نوع الأمر",
                template="streamlit"
            )

            st.plotly_chart(fig, use_container_width=True)


        # *******************************************************************
        # تحويل رؤوس الأعمدة التي تحتوي على تواريخ إلى صيغة مختصرة "يوم شهر"
        # *******************************************************************
        plan_df.columns = [
            col.strftime("%d %b") if isinstance(col, (datetime.datetime, pd.Timestamp)) else col
            for col in plan_df.columns
        ]


        # *******************************************************************
        # زر إنشاء النسخة المضغوطة وحفظها في session_state
        # *******************************************************************
        if st.button("💾 Excel حفظ الملف كـ "):
            current_date = datetime.datetime.now().strftime("%d_%b_%Y")
            
# -------------------------------
            # ✅ تجهيز شيت "BOM_Level1_Expanded" بالمنطق التكراري لاستخراج جميع الفروع
            # -------------------------------
            
            def generate_bom_paths_recursive(bom_df, current_node, current_path, all_paths):
                """
                تستخرج جميع مسارات الـ BOM بشكل تكراري (شاملة الفروع المتعددة).
                """
                
                # إيجاد جميع الأبناء المباشرين للعقدة الحالية (Material)
                children = bom_df[bom_df["Material"] == current_node]

                if children.empty:
                    # إذا لم يكن هناك أبناء، فهذا فرع مكتمل (مادة خام)
                    all_paths.append(current_path)
                    return

                # استدعاء الدالة بشكل تكراري لكل طفل
                for _, child_row in children.iterrows():
                    child_component = child_row["Component"]
                    child_desc = child_row.get("Component Description", "")
                    
                    new_level = (child_component, child_desc)
                    new_path = current_path + [new_level]

                    # منع الحلقات التكرارية (Loop prevention)
                    if child_component not in [p[0] for p in current_path]:
                        generate_bom_paths_recursive(bom_df, child_component, new_path, all_paths)
                    else:
                        # إنهاء المسار عند ظهور حلقة تكرارية
                        all_paths.append(new_path)


            # 1. تحديد المكونات التي تبدأ منها المسارات (المستوى 1)
            # نستخدم Component UoM لأن عمود Hierarchy Level قد لا يكون موجوداً أو موثوقاً به في كل الحالات
            if "Hierarchy Level" in component_df.columns:
                 level1_starters = component_df[component_df["Hierarchy Level"] == 1]
            else:
                 # إذا لم يكن عمود Hierarchy Level موجوداً، نعتبر كل الآباء المحتملين كبداية
                 level1_starters = component_df[['Material', 'Material Description']].drop_duplicates()
                 level1_starters = level1_starters.rename(columns={'Material': 'Component', 'Material Description': 'Component Description'})

            unique_level1_codes = level1_starters["Component"].unique()
            
            rows = []
            max_depth = 0
            
            # 2. توليد المسارات لجميع فروع كل مكون مستوى أول
            for level1_code in unique_level1_codes:
                # نحتاج وصف المكون للمستوى الأول
                desc_row = component_df[component_df["Component"] == level1_code]
                level1_desc = desc_row.iloc[0].get("Component Description", "") if not desc_row.empty else ""
                
                initial_path = [(level1_code, level1_desc)]
                paths_from_starter = []
                
                # بدء الاستدعاء التكراري
                generate_bom_paths_recursive(component_df, level1_code, initial_path, paths_from_starter)
                
                # إضافة المسارات المولدة إلى القائمة الرئيسية
                if paths_from_starter:
                    rows.extend(paths_from_starter)
                else:
                    # إذا كان Level1_Code ليس له أبناء في BOM
                    rows.append(initial_path)
                
                for path in paths_from_starter or [initial_path]:
                    max_depth = max(max_depth, len(path))
            
            # 3. إعداد أعمدة DataFrame
            cols = []
            for i in range(1, max_depth + 1):
                cols.append(f"Level{i}_Code")
                cols.append(f"Level{i}_Desc")

            table_rows = []
            for path in rows:
                row_data = []
                for code, desc in path:
                    row_data.extend([code, desc])
                while len(row_data) < len(cols):
                    row_data.append("") # تعبئة بالخلايا الفارغة
                table_rows.append(row_data)

            bom_levels_df = pd.DataFrame(table_rows, columns=cols)
            # -------------------------------
            # إنشاء ملف Excel وحفظه في session_state
            # -------------------------------
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                plan_df.to_excel(writer, sheet_name="Plan", index=False)
                pivot_by_date.to_excel(writer, sheet_name="Need_By_Date", index=False)
                bom_levels_df.to_excel(writer, sheet_name="BOM_Level1_Expanded", index=False)
                pivot_by_date_multi.to_excel(writer, sheet_name="Need_By_Date Multi level", index=False)
                pivot_by_order.to_excel(writer, sheet_name="Need_By_Order Type", index=False)
                component_bom_pivot.reset_index().to_excel(writer, sheet_name="Component_in_BOMs", index=False)
                component_df.to_excel(writer, sheet_name="Component", index=False)
                if not mrp_df.empty:
                    mrp_df.to_excel(writer, sheet_name="MRP Contor", index=False)
                mrp_result_pivot.to_excel(writer, sheet_name="MRP_Result", index=False)

            excel_buffer.seek(0)
            st.session_state["excel_file"] = excel_buffer
            st.subheader("🔥 Excel أضغط هنا  تحميل ملف كامل")
            st.session_state["excel_file_date"] = current_date  # حفظ التاريخ في session_state

        # زر التحميل باستخدام النسخة المخزنة
        if "excel_file" in st.session_state:
            st.download_button(
                label="📊 تحميل ملف Excel",
                data=st.session_state["excel_file"],
                file_name=f"All_Component_Results_{st.session_state['excel_file_date']}.xlsx",  # استخدام session_state
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("✅ تم إنشاء الملف بنجاح، وجميع الشيتات موجودة داخل Excel")

# --- التذييل ---
st.markdown(
    """
    <hr>
    <div style="text-align:center; direction:rtl; font-size:14px; color:gray;">
        ✨ تم التنفيذ بواسطة <b>م / رضا رشدي</b> – جميع الحقوق محفوظة © 2025 ✨
    </div>
    """,
    unsafe_allow_html=True
)



