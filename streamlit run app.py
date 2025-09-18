
import streamlit as st
import pandas as pd
import datetime
from io import BytesIO
import zipfile
import calendar
import plotly.express as px

# إعداد الصفحة
st.set_page_config(page_title="🔥 MRP Tool", page_icon="📂", layout="wide")
st.subheader("📂 برنامج استخراج وحفظ نتائج الـ MRP")
# صندوق التعريف القابل للطي
#with st.expander("📘 تعريف البرنامج"):
 #   with open("README.md", "r", encoding="utf-8") as f:
  #      readme_content = f.read()
   # st.markdown(readme_content, unsafe_allow_html=True)
st.markdown(
    "<p style='font-size:16px; font-weight:bold;'>📂 اختر ملف الخطة الشهرية Excel</p>",
    unsafe_allow_html=True
)

uploaded_file = st.file_uploader("", type=["xlsx"])

#uploaded_file = st.file_uploader("📂  اختر ملف الخطة الشهرية  Excel", type=["xlsx"])

if uploaded_file:
    with st.spinner("⏳ جاري معالجة البيانات ----- انتظر قليلا.....⏳"):

        # -------------------------------
        # قراءة شيتات Excel
        # -------------------------------
        xls = pd.ExcelFile(uploaded_file, engine='openpyxl')
        plan_df = xls.parse("plan")
        component_df = xls.parse("Component")
        mrp_df = xls.parse("MRP Contor") if "MRP Contor" in xls.sheet_names else pd.DataFrame()
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
        required_component_columns = ["Material", "Component", "Component Quantity"]
        missing_component_columns = [col for col in required_component_columns if col not in component_df.columns]
        if missing_component_columns:
            st.error(f"❌ جدول المكونات لا يحتوي على الأعمدة المطلوبة: {', '.join(missing_component_columns)}")
            st.stop()

        # -------------------------------
        # تجهيز البيانات الأساسية
        # -------------------------------
        plan_melted = plan_df.melt(
            id_vars=["Material", "Material Description", "Order Type"],
            var_name="Date",
            value_name="Planned Quantity"
        )
        plan_melted["Date"] = pd.to_datetime(plan_melted["Date"], errors='coerce')
        merged_df = pd.merge(plan_melted, component_df, on="Material", how="left")
        merged_df["Required Component Quantity"] = merged_df["Planned Quantity"] * merged_df["Component Quantity"]

        # -------------------------------
        # الملخص السريع (عرض فقط)
        # -------------------------------
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

        # -------------------------------
        # Need_By_Date
        # -------------------------------
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

        # -------------------------------
        # Need_By_Order Type
        # -------------------------------
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

        if not mrp_df.empty:
            pivot_by_order = pd.merge(
                pivot_by_order,
                mrp_df[["Component", "MRP Contor"]],
                on="Component",
                how="left"
            )

            # إعادة ترتيب الأعمدة بحيث MRP Contor يكون العمود الثالث
            cols = pivot_by_order.columns.tolist()
            fixed_order = ["Component", "Component Description", "MRP Contor", "Component UoM"]
            other_cols = [c for c in cols if c not in fixed_order]
            pivot_by_order = pivot_by_order[fixed_order + other_cols]

        # تنسيق أسماء الأعمدة (التاريخ + Order Type)
        pivot_by_order.columns = [
            f"{col[1]} - {col[0].strftime('%d %b')}" if isinstance(col, tuple) and isinstance(col[0], pd.Timestamp)
            else col if isinstance(col, str) else col[0]
            for col in pivot_by_order.columns
        ]

         # -------------------------------
         # Component_in_BOMs
         # -------------------------------
#        if not mrp_df.empty:
 #           merged_df = merged_df.merge(mrp_df[["Component", "MRP Contor"]], on="Component", how="left")
         # نعمل Pivot باستخدام Order Type
#        component_bom_map = merged_df.groupby(
 #           ["MRP Contor", "Component", "Material"]
  #      )["Order Type"].apply(lambda x: ','.join(sorted(set(x)))).reset_index(name="Order Type")
   #     component_bom_pivot = component_bom_map.pivot_table(
    #        index=["MRP Contor", "Component"],
     #       columns="Material",
      #      values="Order Type",
       #     aggfunc=lambda x: ','.join(sorted(set(x))),
        #    fill_value="")


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


        # -------------------------------
        # جدول الكميات الشهرية + الرسم البياني
        # -------------------------------
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
            st.subheader("👌 توزيع الكميات الشهرية حسب نوع الأمر")
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


            st.subheader("👌 رسم بياني للكميات الشهرية✅")

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


   	# -------------------------------
        # تحويل رؤوس الأعمدة التي تحتوي على تواريخ إلى صيغة مختصرة "يوم شهر"
        # -------------------------------
        plan_df.columns = [
            col.strftime("%d %b") if isinstance(col, (datetime.datetime, pd.Timestamp)) else col
            for col in plan_df.columns
        ]


        # -------------------------------
        # زر إنشاء النسخة المضغوطة
        # -------------------------------
        if st.button("🗜️ اضغط هنا لإنشاء النسخة المضغوطة"):
            current_date = datetime.datetime.now().strftime("%d_%b_%Y")

            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                plan_df.to_excel(writer, sheet_name="Plan", index=False)
                pivot_by_date.to_excel(writer, sheet_name="Need_By_Date", index=False)
                pivot_by_order.to_excel(writer, sheet_name="Need_By_Order Type", index=False)
                component_bom_pivot.reset_index().to_excel(writer, sheet_name="Component_in_BOMs", index=False)
                component_df.to_excel(writer, sheet_name="Component", index=False)
                if not mrp_df.empty:
                    mrp_df.to_excel(writer, sheet_name="MRP Contor", index=False)
            excel_buffer.seek(0)

            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zipf:
                zipf.writestr(f"All_Component_Results_{current_date}.xlsx", excel_buffer.getvalue())
            zip_buffer.seek(0)

            st.subheader("🔥 تحميل النسخة الكاملة مضغوطة")
            st.download_button(
                label=" 📊 تحميل الملف المضغوط",
                data=zip_buffer,
                file_name=f"All_Component_Results_{current_date}.zip",
                mime="application/zip"
            )

            st.success("✅ تم إنشاء النسخة المضغوطة بنجاح، وجميع الشيتات موجودة داخل Excel")
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




