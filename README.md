# Quinncia PowerPoint Updater

A Streamlit app that updates slide 2, **Appendix: Entire MSB**, in the Career Readiness Milestone Report PowerPoint.

## What it does

- Upload a `.pptx` report/template.
- Upload a Quinncia metrics `.csv` or `.xlsx` export.
- Pull only the first section named `Quinncia Metrics (All Students)`.
- Update slide 2 values in the Appendix table.
- Preserve PowerPoint program labels like `HR`, `BSIS`, `MISM`, and `Overall MSB` while matching them to the Quinncia export rows.
- Make the updated table body text black.
- Leave the rest of the deck unchanged.

## Local setup

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Cloud setup

1. Put these files in a GitHub repository:
   - `app.py`
   - `quinncia_pptx_updater.py`
   - `requirements.txt`
2. Deploy the repo on Streamlit Cloud.
3. Upload the PowerPoint and Quinncia export in the app.
4. Click **Update PowerPoint** and download the finished deck.

## Expected Quinncia source

The CSV can be a multi-section export like this:

```text
Quinncia Metrics (All Students)
major,enrolled_students,quinncia_sign_ups,...
Accounting (BS),311,133,...
```

The app ignores later sections such as class-specific metrics, job search stats, and internship search stats.
