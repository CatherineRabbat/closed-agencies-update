import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

# Streamlit App Title
st.title("Agency Account Update")

# Upload files
closed_agencies_file = st.file_uploader("Upload Closed Agencies File", type=["xlsx"])
accounts_file = st.file_uploader("Upload Salesforce File", type=["xlsx"])

if closed_agencies_file and accounts_file:
    try:
        # Load the data
        closed_agencies = pd.read_excel(closed_agencies_file)
        accounts = pd.read_excel(accounts_file)

        # Match the records based on IATA number
        merged = accounts.merge(closed_agencies, left_on="IATA Number", right_on="IATA Number", how="inner")

        # Update IATA status to "Not Valid"
        merged.loc[:, 'IATA Status'] = "Not Valid"

        # Add "CLOSED-" prefix only if it's not already there
        merged.loc[~merged['Account Name'].str.startswith("CLOSED-", na=False), 'Account Name'] = "CLOSED-" + merged['Account Name']

        # Set Parent Account ID and Ultimate Parent.1
        merged[['Parent Account ID', 'Ultimate Parent.1']] = "001w000001hvcFn"

        # Select required columns
        filtered = merged[['Account ID', 'Account Name', 'IATA Number', 'IATA Status', 'Parent Account ID', 'Ultimate Parent.1']]

        # Identify closed agencies not found in Salesforce
        missing_agencies = closed_agencies[~closed_agencies["IATA Number"].isin(accounts["IATA Number"])]

        st.success("Files processed successfully!")

        # Display the processed data
        st.subheader("Updated Accounts")
        st.dataframe(filtered)

        st.subheader("Missing Agencies")
        st.dataframe(missing_agencies)

        # Provide a download link for the results
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            filtered.to_excel(writer, sheet_name="Updated Accounts", index=False)
            missing_agencies.to_excel(writer, sheet_name="Missing Agencies", index=False)
        output.seek(0)

        st.download_button(label="Download Processed File", data=output, file_name="Updated_Accounts.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"An error occurred: {e}")
