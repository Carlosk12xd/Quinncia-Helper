# Quinncia PowerPoint Updater

Streamlit app that updates Quinncia appendix tables in the Career Readiness Milestone Report PowerPoint.

## What it updates

- Slide 2: **Appendix: Entire MSB** using the CSV/Excel section **Quinncia Metrics (All Students)**
- Slide 4: **Appendix: Class of 2027** using the CSV/Excel section **Quinncia Metrics (Class of 2027 and Above)**

The app preserves the PowerPoint program labels, including HR, BSIS, MISM, and Overall MSB, while matching those rows to the Quinncia export names. It also makes the updated table text black.

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
