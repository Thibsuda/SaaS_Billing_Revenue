# SaaS Billing & Revenue Hub

The problem today: HubSpot knows what was sold and Stripe knows what was billed, but nothing checks them against each other. Each month someone exports both, pastes them into Excel, and rebuilds the same formulas by hand.

This is a finance operations tool for SaaS companies: billing data quality, AR aging, revenue recognition and deferred revenue from HubSpot and Stripe exports and other CRM system.

1. Upload one Excel file with two sheets, **Raw_HubSpot** and **Raw_Stripe**. 
2. Open web : https://saasbillingrevenue-app.streamlit.app/
3. Upload that raw file to web


Pros

-Saves time: hours of Excel work become seconds.
-Same rules every month, so there's less human error.
-Catches duplicates, missing Deal IDs, and sold-but-never-billed deals.
-Built-in tie-out checks, so the team can trust the numbers.
-One shared view for AR, billing and the accountant.
-Downloads an Excel workbook with live formulas for the close package.
-Free to run, and it's easy to change the rules in the code.

Cons

-Works on exported files today, not live HubSpot and Stripe data.
-Tested on invented data. Real data will have more edge cases, like credits, refunds, partial payments and multi-line invoices.
-Revenue is straight-line by day, so usage fees or several performance obligations need extra rules.
-Needs clean exports. Missing or misspelled columns have to be fixed first.
-A person still has to review the output, and the app can't confirm that the exported data is complete.
-Needs someone to maintain it, and a company would need to review security before using real data.

Next steps that fix most cons: connect to the HubSpot and Stripe APIs, add credit and refund rules, and run it on a month of real data.



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
