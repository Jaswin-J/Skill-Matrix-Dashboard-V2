import pandas as pd
import streamlit as st
import os
import seaborn as sns
import matplotlib.pyplot as plt
from matplotlib.colors import rgb2hex

st.set_page_config(page_title="Employee Skill Matrix", layout="wide")

# Function to rename duplicate column names
def rename_duplicate_columns(columns):
    seen = {}
    new_columns = []
    for col in columns:
        if col in seen:
            seen[col] += 1
            new_columns.append(f"{col}_{seen[col]}")
        else:
            seen[col] = 0
            new_columns.append(col)
    return new_columns

# Function to clean score values
def clean_scores(df):
    for col in df.columns[1:]:  # Assuming first column is Employee Names
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        df[col] = df[col].astype(str).apply(lambda x: x.rstrip('0').rstrip('.') if '.' in x else x)
    return df
        

# Function to load data from the Data folder
@st.cache_data
def load_data(file_name):
    file_path = os.path.join(os.getcwd(), file_name)
    raw_df = pd.read_excel(
        file_path,
        #skiprows=1,
        header=[0, 1],
        sheet_name="Employees sheet",
        engine="openpyxl"
    )
    cleaned_columns = rename_duplicate_columns(
        ['_'.join([str(c) for c in col if pd.notna(c)]).strip() for col in raw_df.columns]
    )
    raw_df.columns = cleaned_columns
    cleaned_df = clean_scores(raw_df.copy())
    
    category_mapping = {}
    for col in raw_df.columns[1:]:
        category, subcategory = col.split('_', 1)
        if category not in category_mapping:
            category_mapping[category] = []
        category_mapping[category].append(subcategory)
    
    return raw_df, cleaned_df, category_mapping

st.title("📊 Employee Skill Matrix Dashboard")
st.sidebar.header("⚙️ Select Data File")

# Get list of available Excel files in the Data folder
# data_folder = "Data"
available_files = [f for f in os.listdir(os.getcwd()) if f.endswith(".xlsx")]
selected_file = st.sidebar.selectbox("📂 Select an Excel File", available_files)

if selected_file:
    raw_df, cleaned_df, category_mapping = load_data(selected_file)  # Ensure proper unpacking
    dynamic_categories = list(category_mapping.keys())
    selected_categories = st.sidebar.multiselect("📌 Select Categories", dynamic_categories)

    
    selected_subcategories = {}
    scores = {}
    for category in selected_categories:
        subcategories = category_mapping[category]
        with st.sidebar.expander(f"{category} 🔽"):
            selected_subs = st.multiselect(f"Select Subcategories", subcategories, key=f"subs_{category}")
            if selected_subs:
                selected_subcategories[category] = selected_subs
                for subcat in selected_subs:
                    full_col_name = f"{category}_{subcat}"
                    scores[full_col_name] = st.slider(
                        f"Minimum Score for {subcat}", 1, 5, 3, step=1, key=full_col_name
                    )
    
    filter_type = st.sidebar.radio("Filter Type", ["Match All Conditions", "Match At Least One"], index=0)
    apply_filter_btn = st.sidebar.button("✅ Apply Filters")
    
    if apply_filter_btn and selected_subcategories:
        filtered_df = cleaned_df.copy()
        conditions = []

        for category, subcats in selected_subcategories.items():
            for subcat in subcats:
                full_col_name = f"{category}_{subcat}"
                if full_col_name in filtered_df.columns:
                    filtered_df[full_col_name] = pd.to_numeric(filtered_df[full_col_name], errors='coerce')  # Convert back to numeric
                    conditions.append(filtered_df[full_col_name] >= scores[full_col_name])

        if conditions:
            combined_condition = pd.concat(conditions, axis=1).all(axis=1) if filter_type == "Match All Conditions" else pd.concat(conditions, axis=1).any(axis=1)
            filtered_df = filtered_df.loc[combined_condition]


        if filtered_df.empty:
            st.warning("⚠️ No matching records found!")
        else:
            employee_col = cleaned_df.columns[0]
            filter_cols = [f"{category}_{subcat}" for category, subcats in selected_subcategories.items() for subcat in subcats]
            display_cols = [employee_col] + filter_cols
            display_df = filtered_df[display_cols].copy().reset_index(drop=True)
            display_df.columns = [col.split("_")[-1] if "_" in col else col for col in display_df.columns]
            display_df.index += 1
            
            def color_scale(val):
                colors = sns.color_palette("RdYlGn", n_colors=100)  # More granular gradient
                val_index = int((val / 5) * 99)  # Scale 0-5 to 0-99 index
                background_color = rgb2hex(colors[val_index])
                font_color = "white" if val_index < 40 else "black"  # Ensure contrast
                return f'background-color: {background_color}; color: {font_color}; font-weight: bold;font-size: 16px;'
            
            display_df.iloc[:, 1:] = display_df.iloc[:, 1:].applymap(lambda x: int(x) if x == int(x) else x)  # Remove trailing zeros
            st.write("### 🏆 Filtered Employee Report", unsafe_allow_html=True)
            st.markdown("<div style='text-align: center;'>", unsafe_allow_html=True)
            st.dataframe(display_df.style.applymap(color_scale, subset=display_df.columns[1:]))
            st.markdown("</div>", unsafe_allow_html=True)
            
            csv = filtered_df[display_cols].to_csv(index=False).encode('utf-8')
            st.download_button(
                "📥 Download Filtered Report", data=csv, file_name='filtered_report.csv', mime='text/csv'
            )
    elif apply_filter_btn:
        st.warning("⚠️ Please select at least one category, subcategory, and score filter before generating the report!")
else:
    st.info("📌 Please select an Excel file from the Data folder.")