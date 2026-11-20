# SOC Report PDF to Excel

This Streamlit app wraps the Seceon SOC report converter. Users upload only a
text-based Seceon report PDF. The app uses `SOC_Dashboard_Template.xlsx` stored
next to `app.py`; that Excel template must contain the `Data` and
`Executive Dashboard` sheets. The provided
`SOC_Dashboard_Template.xlsx` is included in this project and used automatically.

## Run locally

```bash
python -m pip install -r requirements.txt
streamlit run app.py
```

Upload the PDF in the app, select **Convert**, review the validation checks,
then download the generated workbook. The uploaded PDF and generated workbook
are processed in memory; no output file is saved to disk.

## Deploy on Streamlit Community Cloud

1. Push `app.py`, `soc_pdf_to_excel.py`, `requirements.txt`, and
   `SOC_Dashboard_Template.xlsx` to a GitHub repository.
2. Sign in at [share.streamlit.io](https://share.streamlit.io/) with GitHub and
   choose **Create app**.
3. Select the repository and branch, set the main file to `app.py`, and deploy.

The template contains the dashboard structure and must stay in the repository
next to `app.py`. Do not commit PDF reports that contain sensitive information.
Streamlit Community Cloud installs dependencies from `requirements.txt`.

The PDF parser is tailored to the report layout and uses its text/vector layer,
not OCR. Check the generated `Checks` sheet and validate the dashboard against
the source report, especially if the report layout has changed.
