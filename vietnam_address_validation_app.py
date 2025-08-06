import streamlit as st
import pandas as pd
import unicodedata
import re
from io import BytesIO

# --- Helper functions ---

def remove_tones(text):
    if not isinstance(text, str):
        return ""
    text = unicodedata.normalize('NFD', text)
    return ''.join(ch for ch in text if unicodedata.category(ch) != 'Mn')

def normalize_text(text):
    if not isinstance(text, str):
        return ""
    return remove_tones(text.strip().lower())

def extract_number_and_street(addr_line1):
    text = normalize_text(addr_line1)
    # Try to extract house number + street name (stop before comma or ward/district/city keywords)
    # This is a heuristic, adjust if needed
    match = re.match(r"(\d+\w*\s[\w\s]+?)(,| ward| district| city|$)", text)
    if match:
        return match.group(1).strip()
    else:
        return text

def flexible_address_match(forms_line1, forms_line2, ups_line1, ups_line2):
    f_ns = extract_number_and_street(forms_line1)
    u_ns = extract_number_and_street(ups_line1)
    f_street = normalize_text(forms_line2)
    u_street = normalize_text(ups_line2)
    return (f_ns == u_ns) and (f_street == u_street)

# --- Core Processing Function ---

def process_files(forms_df, ups_df):
    matched_rows = []
    unmatched_rows = []
    upload_template_rows = []

    # Normalize account number column names for internal processing
    forms_df['Account Number_norm'] = forms_df['Account Number'].astype(str).str.lower().str.strip()
    ups_df['Account Number_norm'] = ups_df['Account Number'].astype(str).str.lower().str.strip()

    ups_groups = ups_df.groupby('Account Number_norm')

    processed_indices = set()

    for idx, form_row in forms_df.iterrows():
        acc_norm = form_row['Account Number_norm']
        same_billing_flag = str(form_row.get("Is Your New Billing Address the Same as Your Pickup and Delivery Address?", "")).strip().lower()

        if acc_norm not in ups_groups.groups:
            # Account number not found in UPS
            unmatched = form_row.to_dict()
            unmatched['Unmatched Reason'] = "Account Number not found in UPS system"
            unmatched_rows.append(unmatched)
            continue

        ups_acc_df = ups_groups.get_group(acc_norm)
        ups_pickup_count = (ups_acc_df['Address Type'] == '02').sum()

        if same_billing_flag == 'yes':
            # Single combined address of type '01' in UPS
            new_addr1 = form_row.get("New Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
            new_addr2 = form_row.get("New Address Line 2 (Street Name)-In English Only", "")
            new_addr3 = form_row.get("New Address Line 3 (Ward/Commune)-In English Only", "")
            city = form_row.get("City / Province", "")
            contact = form_row.get("Full Name of Contact-In English Only", "")

            matched = False
            ups_row_matched = None
            for _, ups_row in ups_acc_df.iterrows():
                if ups_row['Address Type'] == '01':  # All types unified
                    if flexible_address_match(new_addr1, new_addr2, ups_row["Address Line 1"], ups_row["Address Line 2"]):
                        matched = True
                        ups_row_matched = ups_row
                        break

            if matched:
                # Matched entry
                matched_dict = form_row.to_dict()
                matched_dict["New Address Line 1 (Tone-free)"] = normalize_text(new_addr1)
                matched_dict["New Address Line 2 (Tone-free)"] = normalize_text(new_addr2)
                matched_dict["New Address Line 3 (Tone-free)"] = normalize_text(new_addr3)
                matched_rows.append(matched_dict)
                processed_indices.add(idx)

                # Upload template: 3 rows for billing with codes 1, 2, 6
                for code in ['1', '2', '6']:
                    upload_template_rows.append({
                        "AC_NUM": form_row["Account Number"],
                        "AC_Address_Type": code,
                        "invoice option": code,
                        "AC_Name": ups_row_matched["AC_Name"],
                        "Address_Line1": normalize_text(new_addr1),
                        "Address_Line2": normalize_text(new_addr2),
                        "City": city,
                        "Postal_Code": ups_row_matched["Postal_Code"],
                        "Country_Code": ups_row_matched["Country_Code"],
                        "Attention_Name": contact,
                        "Address_Line22": normalize_text(new_addr3),
                        "Address_Country_Code": ups_row_matched["Address_Country_Code"]
                    })
            else:
                unmatched = form_row.to_dict()
                unmatched['Unmatched Reason'] = "Unified billing address not matched in UPS system"
                unmatched_rows.append(unmatched)

        else:
            # Billing, delivery, pickup addresses are separate

            billing_addr1 = form_row.get("New Billing Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
            billing_addr2 = form_row.get("New Billing Address Line 2 (Street Name)-In English Only", "")
            billing_addr3 = form_row.get("New Billing Address Line 3 (Ward/Commune)-In English Only", "")
            billing_city = form_row.get("New Billing City / Province", "")

            delivery_addr1 = form_row.get("New Delivery Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
            delivery_addr2 = form_row.get("New Delivery Address Line 2 (Street Name)-In English Only", "")
            delivery_addr3 = form_row.get("New Delivery Address Line 3 (Ward/Commune)-In English Only", "")
            delivery_city = form_row.get("New Delivery City / Province", "")

            # Pickup addresses count and pickup addresses from form
            pickup_count_form = 0
            try:
                pickup_count_form = int(form_row.get("How Many Pick Up Address Do You Have?", 0))
            except:
                pickup_count_form = 0
            pickup_count_form = min(pickup_count_form, 3)

            pickup_addrs = []
            for i in range(1, pickup_count_form + 1):
                prefix = ["First", "Second", "Third"][i - 1] + " New Pick Up Address"
                pu_addr1 = form_row.get(f"{prefix} Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
                pu_addr2 = form_row.get(f"{prefix} Line 2 (Street Name)-In English Only", "")
                pu_addr3 = form_row.get(f"{prefix} Line 3 (Ward/Commune)-In English Only", "")
                pu_city = form_row.get(f"{prefix} City / Province", "")
                pickup_addrs.append((pu_addr1, pu_addr2, pu_addr3, pu_city))

            def find_match(addr1, addr2, addr_type):
                for _, ups_row in ups_acc_df.iterrows():
                    if ups_row["Address Type"] == addr_type:
                        if flexible_address_match(addr1, addr2, ups_row["Address Line 1"], ups_row["Address Line 2"]):
                            return ups_row
                return None

            billing_match = find_match(billing_addr1, billing_addr2, "03")
            delivery_match = find_match(delivery_addr1, delivery_addr2, "13")

            # Validate pickup address count match
            if pickup_count_form != ups_pickup_count:
                unmatched = form_row.to_dict()
                unmatched['Unmatched Reason'] = f"Pickup address count mismatch - Forms: {pickup_count_form}, UPS: {ups_pickup_count}"
                unmatched_rows.append(unmatched)
                continue

            # Validate each pickup address
            pickup_matches = []
            all_pickup_matched = True
            for pu_addr in pickup_addrs:
                pu_match = find_match(pu_addr[0], pu_addr[1], "02")
                if pu_match is None:
                    unmatched = form_row.to_dict()
                    unmatched['Unmatched Reason'] = f"Pickup address not matched: {pu_addr[0]}, {pu_addr[1]}"
                    unmatched_rows.append(unmatched)
                    all_pickup_matched = False
                    break
                else:
                    pickup_matches.append(pu_match)

            if all_pickup_matched and billing_match and delivery_match:
                processed_indices.add(idx)
                matched = form_row.to_dict()
                # Remove tones in matched file
                matched["New Billing Address Line 1 (Tone-free)"] = normalize_text(billing_addr1)
                matched["New Billing Address Line 2 (Tone-free)"] = normalize_text(billing_addr2)
                matched["New Billing Address Line 3 (Tone-free)"] = normalize_text(billing_addr3)
                matched["New Delivery Address Line 1 (Tone-free)"] = normalize_text(delivery_addr1)
                matched["New Delivery Address Line 2 (Tone-free)"] = normalize_text(delivery_addr2)
                matched["New Delivery Address Line 3 (Tone-free)"] = normalize_text(delivery_addr3)

                for i, pu_addr in enumerate(pickup_addrs, 1):
                    matched[f"{['First', 'Second', 'Third'][i-1]} New Pick Up Address Line 1 (Tone-free)"] = normalize_text(pu_addr[0])
                    matched[f"{['First', 'Second', 'Third'][i-1]} New Pick Up Address Line 2 (Tone-free)"] = normalize_text(pu_addr[1])
                    matched[f"{['First', 'Second', 'Third'][i-1]} New Pick Up Address Line 3 (Tone-free)"] = normalize_text(pu_addr[2])

                matched_rows.append(matched)

                # Pickup addresses rows in upload template - invoice option empty
                for pu_addr in pickup_addrs:
                    upload_template_rows.append({
                        "AC_NUM": form_row["Account Number"],
                        "AC_Address_Type": "02",
                        "invoice option": "",
                        "AC_Name": ups_acc_df["AC_Name"].values[0],
                        "Address_Line1": normalize_text(pu_addr[0]),
                        "Address_Line2": normalize_text(pu_addr[1]),
                        "City": pu_addr[3],
                        "Postal_Code": ups_acc_df["Postal_Code"].values[0],
                        "Country_Code": ups_acc_df["Country_Code"].values[0],
                        "Attention_Name": form_row.get("Full Name of Contact-In English Only", ""),
                        "Address_Line22": normalize_text(pu_addr[2]),
                        "Address_Country_Code": ups_acc_df["Address_Country_Code"].values[0]
                    })

                # Billing address rows with 3 codes & invoice options 1, 2, 6
                for code in ['1', '2', '6']:
                    upload_template_rows.append({
                        "AC_NUM": form_row["Account Number"],
                        "AC_Address_Type": code,
                        "invoice option": code,
                        "AC_Name": ups_acc_df["AC_Name"].values[0],
                        "Address_Line1": normalize_text(billing_addr1),
                        "Address_Line2": normalize_text(billing_addr2),
                        "City": billing_city,
                        "Postal_Code": ups_acc_df["Postal_Code"].values[0],
                        "Country_Code": ups_acc_df["Country_Code"].values[0],
                        "Attention_Name": form_row.get("Full Name of Contact-In English Only", ""),
                        "Address_Line22": normalize_text(billing_addr3),
                        "Address_Country_Code": ups_acc_df["Address_Country_Code"].values[0]
                    })

                # Delivery address single row, invoice option empty
                upload_template_rows.append({
                    "AC_NUM": form_row["Account Number"],
                    "AC_Address_Type": "13",
                    "invoice option": "",
                    "AC_Name": ups_acc_df["AC_Name"].values[0],
                    "Address_Line1": normalize_text(delivery_addr1),
                    "Address_Line2": normalize_text(delivery_addr2),
                    "City": delivery_city,
                    "Postal_Code": ups_acc_df["Postal_Code"].values[0],
                    "Country_Code": ups_acc_df["Country_Code"].values[0],
                    "Attention_Name": form_row.get("Full Name of Contact-In English Only", ""),
                    "Address_Line22": normalize_text(delivery_addr3),
                    "Address_Country_Code": ups_acc_df["Address_Country_Code"].values[0]
                })

            else:
                # Not all matched
                # If billing or delivery not matched, add reason
                if billing_match is None:
                    unmatched = form_row.to_dict()
                    unmatched['Unmatched Reason'] = "Billing address not matched in UPS system"
                    unmatched_rows.append(unmatched)
                    continue
                if delivery_match is None:
                    unmatched = form_row.to_dict()
                    unmatched['Unmatched Reason'] = "Delivery address not matched in UPS system"
                    unmatched_rows.append(unmatched)
                    continue
                if not all_pickup_matched:
                    # Already appended unmatched inside loop
                    continue

    # Add any unprocessed rows as unmatched with generic reason
    unmatched_not_processed = forms_df.loc[~forms_df.index.isin(processed_indices)]
    for _, row in unmatched_not_processed.iterrows():
        unmatched = row.to_dict()
        unmatched['Unmatched Reason'] = "No matching address found or not processed"
        unmatched_rows.append(unmatched)

    matched_df = pd.DataFrame(matched_rows)
    unmatched_df = pd.DataFrame(unmatched_rows)
    upload_template_df = pd.DataFrame(upload_template_rows)

    return matched_df, unmatched_df, upload_template_df

# --- Streamlit UI ---
def main():
    st.set_page_config(page_title="Vietnam Address Validation Tool", layout="wide")
    st.title("🇻🇳 Vietnam Address Validation Tool")
    st.write("Upload Microsoft Forms response file and UPS system address file to validate and generate upload template.")

    forms_file = st.file_uploader("Upload Microsoft Forms Response File (.xlsx)", type=["xlsx"])
    ups_file = st.file_uploader("Upload UPS System Address File (.xlsx)", type=["xlsx"])

    if forms_file and ups_file:
        with st.spinner("Reading files..."):
            forms_df = pd.read_excel(forms_file)
            ups_df = pd.read_excel(ups_file)

        # Required columns check
        required_forms_cols = [
            "Account Number",
            "Is Your New Billing Address the Same as Your Pickup and Delivery Address?"
        ]
        required_ups_cols = [
            "Account Number",
            "Address Type",
            "Address Line 1",
            "Address Line 2",
            "AC_Name",
            "Postal_Code",
            "Country_Code",
            "Address_Country_Code"
        ]

        for col in required_forms_cols:
            if col not in forms_df.columns:
                st.error(f"Missing column in Forms file: {col}")
                return

        for col in required_ups_cols:
            if col not in ups_df.columns:
                st.error(f"Missing column in UPS file: {col}")
                return

        matched_df, unmatched_df, upload_template_df = process_files(forms_df, ups_df)

        st.success(f"Validation completed! Matched: {len(matched_df)}, Unmatched: {len(unmatched_df)}")

        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')
            return output.getvalue()

        if not matched_df.empty:
            st.download_button(
                label="Download Matched Records",
                data=to_excel(matched_df),
                file_name="matched_records.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        if not unmatched_df.empty:
            st.download_button(
                label="Download Unmatched Records",
                data=to_excel(unmatched_df),
                file_name="unmatched_records.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        if not upload_template_df.empty:
            st.download_button(
                label="Download Upload Template",
                data=to_excel(upload_template_df),
                file_name="upload_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
