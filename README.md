# Quinncia PowerPoint Updater

Streamlit app that updates Quinncia appendix tables in the Career Readiness Milestone Report PowerPoint.

## What it updates

- Slide 1: **Career Launch Readiness KPIs** using the Overall Marriott School row from **Quinncia Metrics (All Students)**
- Slide 2: **Appendix: Entire MSB** using **Quinncia Metrics (All Students)**
- Slide 3: **Career Launch Readiness KPIs: Class of 2027** using the Overall Marriott School row from **Quinncia Metrics (Class of 2027 and Above)**
- Slide 4: **Appendix: Class of 2027** using **Quinncia Metrics (Class of 2027 and Above)**

The app preserves the PowerPoint program labels, including HR, BSIS, MISM, and Overall MSB, while matching those rows to the Quinncia export names. It also updates the date to the day the app runs and makes the updated text black.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy to Streamlit Cloud

1. Upload these files to a GitHub repository:
   - `app.py`
   - `quinncia_pptx_updater.py`
   - `requirements.txt`
   - `README.md`
2. Create a new Streamlit Cloud app.
3. Set the main file path to `app.py`.
4. Upload your PowerPoint and Quinncia CSV/Excel export in the app.
5. Download the updated PowerPoint.
