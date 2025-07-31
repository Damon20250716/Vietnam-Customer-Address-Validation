import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Vietnam Address Validation Tool", layout="wide")
st.title("🇻🇳 Vietnam Address Validation Tool")

def load_excel(file):
    return pd.read_excel(file)

def normalize_text(s):
    return str(s).strip().lower() if pd.notna(s) else ""

def prepare_upload_rows(account, line1, line2, line3, code_list):
    rows = []
    for code in code_list:
        rows.append({
            "Account Number": account,
            "Address Type Code": code,
            "New Address Line 1": line1,
            "New Address Line 2": line2,
            "New Address Line 3": line3
        })
    return rows

def validate_addresses(forms_df, ups_df):
    matched = []
    unmatched = []
    upload_rows = []

    # Normalize for comparison
    ups_df["Account Number"] = ups_df["Account Number"].astype(str).str.strip().str.lower()
    forms_df["Account Number"] = forms_df["Account Number"].astype(str).str.strip().str.lower()

    ups_accounts = set(ups_df["Account Number"])

    for _, row in forms_df.iterrows():
        account = normalize_text(row["Account Number"])
        if account not in ups_accounts:
            row["Issue"] = "Account not found in UPS system"
            unmatched.append(row)
            continue

        use_all = normalize_text(row.get("Is Your New Billing Address the Same as Your Pickup and Delivery Address?")) == "yes"

        if use_all:
            line1 = row.get("New Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
            line2 = row.get("New Address Line 2 (Street Name)-In English Only", "")
            line3 = row.get("New Address Line 3 (Ward/Commune)-In English Only", "")
            if any(pd.notna(v) and str(v).strip() != "" for v in [line1, line2, line3]):
                matched.append({
                    "Account Number": account,
                    "Type": "01",
                    "Line 1": line1,
                    "Line 2": line2,
                    "Line 3": line3
                })
                upload_rows.extend(prepare_upload_rows(account, str(line1).strip(), str(line2).strip(), str(line3).strip(), [1, 2, 6]))
            else:
                row["Issue"] = "No valid unified address data"
                unmatched.append(row)
            continue

        # “No” case – process individual address blocks
        valid = False

        # Billing Address → 3 rows
        b1 = row.get("New Billing Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
        b2 = row.get("New Billing Address Line 2 (Street Name)-In English Only", "")
        b3 = row.get("New Billing Address Line 3 (Ward/Commune)-In English Only", "")
        if any(pd.notna(v) and str(v).strip() != "" for v in [b1, b2, b3]):
            matched.append({"Account Number": account, "Type": "03", "Line 1": b1, "Line 2": b2, "Line 3": b3})
            upload_rows.extend(prepare_upload_rows(account, str(b1).strip(), str(b2).strip(), str(b3).strip(), [1, 2, 6]))
            valid = True

        # Delivery Address → 1 row
        d1 = row.get("New Delivery Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
        d2 = row.get("New Delivery Address Line 2 (Street Name)-In English Only", "")
        d3 = row.get("New Delivery Address Line 3 (Ward/Commune)-In English Only", "")
        if any(pd.notna(v) and str(v).strip() != "" for v in [d1, d2, d3]):
            matched.append({"Account Number": account, "Type": "13", "Line 1": d1, "Line 2": d2, "Line 3": d3})
            upload_rows.extend(prepare_upload_rows(account, str(d1).strip(), str(d2).strip(), str(d3).strip(), [5]))
            valid = True

        # Pickup Addresses → up to 3 rows
        for i in ["First", "Second", "Third"]:
            p1 = row.get(f"{i} New Pick Up Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
            p2 = row.get(f"{i} New Pick Up Address Line 2 (Street Name)-In English Only", "")
            p3 = row.get(f"{i} New Pick Up Address Line 3 (Ward/Commune)-In English Only", "")
            if any(pd.notna(v) and str(v).strip() != "" for v in [p1, p2, p3]):
                matched.append({"Account Number": account, "Type": "02", "Line 1": p1, "Line 2": p2, "Line 3": p3})
                upload_rows.extend(prepare_upload_rows(account, str(p1).strip(), str(p2).strip(), str(p3).strip(), [4]))
                valid = True

        if not valid:
            row["Issue"] = "No valid address block filled"
            unmatched.append(row)

    matched_df = pd.DataFrame(matched)
    unmatched_df = pd.DataFrame(unmatched)
    upload_df = pd.DataFrame(upload_rows)

    return matched_df, unmatched_df, upload_df

def download_excel(df, filename):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    st.download_button(label=f"📥 Download {filename}", data=output, file_name=filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# UI
st.markdown("### 1. Upload Microsoft Forms Response File")
forms_file = st.file_uploader("Upload Forms Excel file", type=["xlsx"])

st.markdown("### 2. Upload UPS Existing System File")
ups_file = st.file_uploader("Upload UPS System Excel file", type=["xlsx"])

if forms_file and ups_file:
    try:
        forms_df = load_excel(forms_file)
        ups_df = load_excel(ups_file)

        st.success("✅ Files loaded successfully.")

        matched_df, unmatched_df, upload_df = validate_addresses(forms_df, ups_df)

        st.markdown("### ✅ Matched Records")
        st.dataframe(matched_df)
        download_excel(matched_df, "matched_records.xlsx")

        st.markdown("### ⚠️ Unmatched Records (Forms Only)")
        st.dataframe(unmatched_df)
        download_excel(unmatched_df, "unmatched_forms_only.xlsx")

        st.markdown("### 📦 Upload Template Format")
        st.dataframe(upload_df)
        download_excel(upload_df, "upload_template.xlsx")

    except Exception as e:
        st.error(f"❌ Error during processing: {e}")
