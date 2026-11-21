# SaaS Billing & Revenue Hub

A finance operations tool for SaaS companies: billing data quality, AR aging, revenue recognition and deferred revenue from HubSpot and Stripe exports and other CRM system.

1. Upload one Excel file with two sheets, **Raw_HubSpot** and **Raw_Stripe**. 
2. Open web : https://saasbillingrevenue-app.streamlit.app/
3. Upload that raw file to web


The app rebuilds every tab of the the workbook workbook: Summary_Dashboard, Dashboard, RevRec_Waterfall, Deferred_Rev, HubSpot_Working, Stripe_Working and Support. You can download the full Excel workbook with live formulas.



## Input
- **Raw_HubSpot:** Deal ID, Customer, Stage, Start Date, Contract Value, End Date, Product
- **Raw_Stripe:** Invoice ID, Customer, Invoice Date, Due Date, Amount, Amount Paid, Status, Deal ID

A sample file is included: `Sample_Raw_Data.xlsx`.

## Files
- `streamlit_app.py` the web app
- `saas_billing_logic.py` all the calculation logic (same rules as the workbook formulas)
- `excel_export.py` builds the downloadable workbook with formulas
- `requirements.txt`, `.streamlit/config.toml`

## Run locally
    pip install -r requirements.txt
    streamlit run streamlit_app.py

## Deploy on Streamlit Community Cloud
1. Create a GitHub repo (for example `SaaS_Billing_Revenue_Hub`) and upload all these files.
2. Go to share.streamlit.io, sign in with GitHub, click **Create app**.
3. Pick the repo, branch `main`, main file `streamlit_app.py`.
4. In **App URL**, type `saas-billing-revenue` (an ampersand is not allowed in a URL).
5. Click **Deploy**.
