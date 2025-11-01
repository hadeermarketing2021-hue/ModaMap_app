import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import load_workbook

FILE_PATH = "ModaMap_Management.xlsx"

# === دوال قراءة وحفظ ===
def load_sheet(sheet_name):
    df = pd.read_excel(FILE_PATH, sheet_name=sheet_name)
    return df

def save_sheet(df, sheet_name):
    with pd.ExcelWriter(FILE_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)

# === Sidebar للاختيار بين الصفحات ===
st.sidebar.title("Moda Map Management")
sheet_choice = st.sidebar.selectbox("Choose Page / اختر الصفحة", 
                                    ["Sales", "Inventory", "Customers", "Feedback", "Dashboard"])

# === التعامل مع الصفحات ===
if sheet_choice != "Dashboard":
    df = load_sheet(sheet_choice)
    st.title(f"{sheet_choice} Page / صفحة {sheet_choice}")
    st.dataframe(df)

    # Form لإضافة صف جديد
    with st.form("Add new row"):
        new_row = {}
        for col in df.columns:
            new_row[col] = st.text_input(f"{col}")
        submitted = st.form_submit_button("Add / أضف")
        if submitted:
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            save_sheet(df, sheet_choice)
            st.success("Added successfully! / تم الإضافة بنجاح")

else:
    st.title("Dashboard / لوحة المتابعة")

    # === Sales metrics ===
    sales = load_sheet("Sales")
    sales["Price / السعر"] = pd.to_numeric(sales["Price / السعر"], errors='coerce')
    total_sales = sales["Price / السعر"].sum()
    st.metric("Total Sales / إجمالي المبيعات", total_sales)

    orders_count = len(sales)
    st.metric("Orders Count / عدد الطلبات", orders_count)

    # === Feedback metrics ===
    feedback = load_sheet("Feedback")
    feedback["Rating (1–5) / التقييم (١–٥)"] = pd.to_numeric(feedback["Rating (1–5) / التقييم (١–٥)"], errors='coerce')
    avg_rating = feedback["Rating (1–5) / التقييم (١–٥)"].mean()
    st.metric("Average Rating / متوسط التقييم", round(avg_rating,2))

    returns_count = feedback["Return? / هل في مرتجع"].apply(lambda x: 1 if str(x).lower()=="yes" else 0).sum()
    st.metric("Return Count / عدد المرتجعات", returns_count)

    st.subheader("Latest Orders / آخر الطلبات")
    st.dataframe(sales.tail(5))

    # === Burnout Chart ===
    st.subheader("Monthly Burnout Chart / رسم متابعة التارجت الشهري")
    if "Date / التاريخ" in sales.columns and not sales.empty:
        sales["Date / التاريخ"] = pd.to_datetime(sales["Date / التاريخ"], errors='coerce')
        monthly_sales = sales.groupby(sales["Date / التاريخ"].dt.to_period("M"))["Price / السعر"].sum()
        fig, ax = plt.subplots()
        monthly_sales.plot(kind='bar', ax=ax, color='skyblue')
        ax.set_ylabel("Total Sales / إجمالي المبيعات")
        ax.set_xlabel("Month / الشهر")
        st.pyplot(fig)
    else:
        st.info("No sales data to show chart / لا توجد بيانات مبيعات للرسم البياني")
