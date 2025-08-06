import streamlit as st
import pandas as pd
from io import BytesIO
import unicodedata

# Remove Vietnamese tones and normalize string: lowercase, strip spaces
def normalize_text(text):
    if not isinstance(text, str):
        return ""
    text = unicodedata.normalize('NFD', text)
    text = ''.join(ch for ch in text if unicodedata.category(ch) != 'Mn')
    text = text.lower().strip()
    return text

# Extract number + street part from Address Line 1 (before first comma or entire line)
def extract_number_street(addr_line1):
    addr_line1 = normalize_text(addr_line1)
    if ',' in addr_line1:
        return addr_line1.split(',')[0].strip()
    else:
        return addr_line1

# Extract number + street part from Address Line 2 (street name)
def extract_street_name(addr_line2):
    return normalize_text(addr_line2)

# Check if number + street match (line1 and line2) between Forms and UPS addresses
def is_address_match(forms_line1, forms_line2, ups_line1, ups_line2):
    return (extract_number_street(forms_line1) == extract_number_street(ups_line1)) and \
           (extract_street_name(forms_line2) == extract_street_name(ups_line2))

def process_files(forms_df, ups_df):
    matched_rows = []
    unmatched_rows = []
    upload_template_rows = []

    # Normalize account number for matching (lowercase & strip)
    ups_df['Account Number_norm'] = ups_df['Account Number'].astype(str).str.lower().str.strip()
    forms_df['Account Number_norm'] = forms_df['Account Number'].astype(str).str.lower().str.strip()

    ups_grouped = ups_df.groupby('Account Number_norm')

    processed_indices = set()

    for idx, form_row in forms_df.iterrows():
        acc_norm = form_row['Account Number_norm']
        is_same_billing = str(form_row.get("Is Your New Billing Address the Same as Your Pickup and Delivery Address?", "")).strip().lower()

        if acc_norm not in ups_grouped.groups:
            unmatched = form_row.to_dict()
            unmatched['Unmatched Reason'] = "Account Number not found in UPS system"
            unmatched_rows.append(unmatched)
            continue

        ups_acc_df = ups_grouped.get_group(acc_norm)

        ups_pickup_count = (ups_acc_df['Address Type'] == '02').sum()

        # Handle "Yes" case — unified billing address, UPS Address Type 01
        if is_same_billing == "yes":
            # Get unified new address lines from forms
            new_addr1 = form_row.get("New Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
            new_addr2 = form_row.get("New Address Line 2 (Street Name)-In English Only", "")
            new_addr3 = form_row.get("New Address Line 3 (Ward/Commune)-In English Only", "")
            city = form_row.get("City / Province", "")
            contact = form_row.get("Full Name of Contact-In English Only", "")

            # Find matching UPS record with Address Type == '01'
            ups_match = None
            for _, ups_row in ups_acc_df.iterrows():
                if ups_row['Address Type'] == '01':
                    if is_address_match(new_addr1, new_addr2, ups_row['Address Line 1'], ups_row['Address Line 2']):
                        ups_match = ups_row
                        break

            if ups_match is None:
                unmatched = form_row.to_dict()
                unmatched['Unmatched Reason'] = "Unified billing address not matched in UPS system"
                unmatched_rows.append(unmatched)
                continue

            # Matched
            processed_indices.add(idx)

            # Add tone-free normalized fields to matched row output
            matched_entry = form_row.to_dict()
            matched_entry["New Address Line 1 (Tone-free)"] = normalize_text(new_addr1)
            matched_entry["New Address Line 2 (Tone-free)"] = normalize_text(new_addr2)
            matched_entry["New Address Line 3 (Tone-free)"] = normalize_text(new_addr3)
            matched_rows.append(matched_entry)

            # Upload template: 1 row with Address Type = 01, invoice option empty (for unified)
            upload_template_rows.append({
                "AC_NUM": form_row["Account Number"],
                "AC_Address_Type": "01",
                "invoice option": "",  # no invoice option for unified address
                "AC_Name": ups_match["AC_Name"],
                "Address_Line1": normalize_text(new_addr1),
                "Address_Line2": normalize_text(new_addr2),
                "City": city,
                "Postal_Code": ups_match["Postal_Code"],
                "Country_Code": ups_match["Country_Code"],
                "Attention_Name": contact,
                "Address_Line22": normalize_text(new_addr3),
                "Address_Country_Code": ups_match["Address_Country_Code"]
            })

        else:
            # "No" case — separate billing, delivery, pickups

            # Extract new billing, delivery, pickup info from forms
            billing_addr1 = form_row.get("New Billing Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
            billing_addr2 = form_row.get("New Billing Address Line 2 (Street Name)-In English Only", "")
            billing_addr3 = form_row.get("New Billing Address Line 3 (Ward/Commune)-In English Only", "")
            billing_city = form_row.get("New Billing City / Province", "")

            delivery_addr1 = form_row.get("New Delivery Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
            delivery_addr2 = form_row.get("New Delivery Address Line 2 (Street Name)-In English Only", "")
            delivery_addr3 = form_row.get("New Delivery Address Line 3 (Ward/Commune)-In English Only", "")
            delivery_city = form_row.get("New Delivery City / Province", "")

            # Number of pickups customer entered
            try:
                pickup_count = int(form_row.get("How Many Pick Up Address Do You Have?", 0))
            except:
                pickup_count = 0
            if pickup_count > 3:
                pickup_count = 3

            # Collect pickup addresses from forms (up to 3)
            pickup_addrs = []
            for i in range(1, pickup_count + 1):
                prefix = ["First", "Second", "Third"][i-1] + " New Pick Up Address"
                pu_addr1 = form_row.get(f"{prefix} Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
                pu_addr2 = form_row.get(f"{prefix} Line 2 (Street Name)-In English Only", "")
                pu_addr3 = form_row.get(f"{prefix} Line 3 (Ward/Commune)-In English Only", "")
                pu_city = form_row.get(f"{prefix} City / Province", "")
                pickup_addrs.append((pu_addr1, pu_addr2, pu_addr3, pu_city))

            # Helper to find UPS row matching by address type and address match
            def find_ups_match(addr1, addr2, addr_type):
                for _, ups_row in ups_acc_df.iterrows():
                    if ups_row['Address Type'] == addr_type:
                        if is_address_match(addr1, addr2, ups_row['Address Line 1'], ups_row['Address Line 2']):
                            return ups_row
                return None

            billing_match = find_ups_match(billing_addr1, billing_addr2, '03')
            delivery_match = find_ups_match(delivery_addr1, delivery_addr2, '13')

            # Validate pickup count matches UPS data exactly
            if pickup_count != ups_pickup_count:
                unmatched = form_row.to_dict()
                unmatched['Unmatched Reason'] = f"Pickup address count mismatch: Forms={pickup_count}, UPS={ups_pickup_count}"
                unmatched_rows.append(unmatched)
                continue

            # Validate each pickup address matches
            pickup_matches = []
            all_pickup_matched = True
            for pu_addr in pickup_addrs:
                pu_match = find_ups_match(pu_addr[0], pu_addr[1], '02')
                if pu_match is None:
                    unmatched = form_row.to_dict()
                    unmatched['Unmatched Reason'] = f"Pickup address not matched: {pu_addr[0]}, {pu_addr[1]}"
                    unmatched_rows.append(unmatched)
                    all_pickup_matched = False
                    break
                else:
                    pickup_matches.append(pu_match)
            if not all_pickup_matched:
                continue

            # Check billing and delivery address matches
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

            # All matched -> add matched record with tone-free normalized fields
            processed_indices.add(idx)
            matched_entry = form_row.to_dict()

            matched_entry["New Billing Address Line 1 (Tone-free)"] = normalize_text(billing_addr1)
            matched_entry["New Billing Address Line 2 (Tone-free)"] = normalize_text(billing_addr2)
            matched_entry["New Billing Address Line 3 (Tone-free)"] = normalize_text(billing_addr3)

            matched_entry["New Delivery Address Line 1 (Tone-free)"] = normalize_text(delivery_addr1)
            matched_entry["New Delivery Address Line 2 (Tone-free)"] = normalize_text(delivery_addr2)
            matched_entry["New Delivery Address Line 3 (Tone-free)"] = normalize_text(delivery_addr3)

            for i, pu_addr in enumerate(pickup_addrs, start=1):
                matched_entry[f"{['First', 'Second', 'Third'][i-1]} New Pick Up Address Line 1 (Tone-free)"] = normalize_text(pu_addr[0])
                matched_entry[f"{['First', 'Second', 'Third'][i-1]} New Pick Up Address Line 2 (Tone-free)"] = normalize_text(pu_addr[1])
                matched_entry[f"{['First', 'Second', 'Third'][i-1]} New Pick Up Address Line 3 (Tone-free)"] = normalize_text(pu_addr[2])

            matched_rows.append(matched_entry)

            # Build upload template rows:

            # 1) Pickups as separate rows (Address Type = 02, invoice option empty)
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

            # 2) Billing address split into 3 rows with codes 1,2,6 and invoice option same as code
            for code in ["1", "2", "6"]:
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

            # 3) Delivery address row (Address Type = 13, invoice option empty)
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

    # Any Forms rows not processed are unmatched
    unmatched_not_processed = forms_df.loc[~forms_df.index.isin(processed_indices)]
    for _, row in unmatched_not_processed.iterrows():
        unmatched = row.to_dict()
        unmatched['Unmatched Reason'] = "No matching address found or not processed"
        unmatched_rows.append(unmatched)

    matched_df = pd.DataFrame(matched_rows)
    unmatched_df = pd.DataFrame(unmatched_rows)
    upload_template_df = pd.DataFrame(upload_template_rows)

    return matched_df, unmatched_df, upload_template_df

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

        # Validate required columns exist
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

        missing_forms = [col for col in required_forms_cols if col not in forms_df.columns]
        missing_ups = [col for col in required_ups_cols if col not in ups_df.columns]

        if missing_forms:
            st.error(f"Missing column(s) in Forms file: {', '.join(missing_forms)}")
            return
        if missing_ups:
            st.error(f"Missing column(s) in UPS file: {', '.join(missing_ups)}")
            return

        matched_df, unmatched_df, upload_template_df = process_files(forms_df, ups_df)

        st.success(f"Validation completed! Matched: {len(matched_df)}, Unmatched: {len(unmatched_df)}")

        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
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
