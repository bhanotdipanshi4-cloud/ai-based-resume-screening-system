# Talent Lens — AI-assisted resume screening demo

A local-first Streamlit app for reviewing multiple PDF and DOCX resumes against a job description. It includes a hiring dashboard, transparent weighted resume scores, skill/experience/education matching, an ATS text-readiness checklist, improvement suggestions, candidate filters/ranking/comparison, resume previews, reviewer-managed statuses, charts, activity, and CSV/PDF exports.

## Run on Windows

1. Open this folder in VS Code and open a terminal in it.
2. Create and activate a virtual environment (recommended):
   - `py -m venv .venv`
   - `\.venv\Scripts\Activate.ps1`
3. Install dependencies: `pip install -r requirements.txt`
4. Start the app: `streamlit run app.py`
5. Open the local URL printed by Streamlit.

## Use

- Paste the role description and optionally add/correct required skills in the sidebar.
- Upload multiple PDF/DOCX resumes, then select **Analyze new resumes**.
- Review matching and missing skills, extracted resume text, and the original PDF preview.
- Inspect the ATS text-readiness checks and optional resume-improvement suggestions for each candidate.
- Apply filters, compare candidates, set a reviewer status, and export the screening report as CSV.
- Export a PDF analysis report from the Activity & export tab.

## Important limitations

- This demo uses deterministic, local text extraction and a transparent skill/alias vocabulary; it does not call an external generative-AI service. The “AI-assisted” label describes lightweight text matching, not a trained model.
- Match percentage is the proportion of selected required skills found in extracted resume text. It is not a measure of overall candidate quality, experience, or suitability. Ranking is only a way to sort that narrow signal.
- The overall resume score is a transparent weighted combination of skill overlap (70%), stated experience match (20%), and education match (10%). Experience and education are included only when a requirement is detected in the job description; criteria extraction is approximate and must be checked. These weights are an example, not a validated hiring model.
- The ATS check is a text-only heuristic for extractable content, contact information, section labels, and bullets. It cannot verify layout compatibility with a specific applicant-tracking system.
- Improvement suggestions are simple text checks, not generative AI or career advice. Candidates should use only truthful and relevant suggestions.
- No shortlist/reject status is assigned automatically. Recommendations are advisory; a human must review the original resume and make every employment decision.
- Scanned/image-only PDFs need OCR before upload. Extraction may miss content in complex layouts; verify any extracted details in the original file.
- Resumes remain in Streamlit session memory and are not written to a database. This is a demo, not a production applicant tracking system. For real candidate data, deploy only in an organization-approved, access-controlled environment and follow privacy, retention, accessibility, and employment requirements.

All project files are in this one folder; there are no generated data folders or database dependencies.
