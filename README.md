# Star Builders - site expense book (Streamlit)

## Run it in VS Code
1. Open this folder in VS Code (File > Open Folder).
2. Terminal > New Terminal, then:
       python -m venv .venv
       .venv\Scripts\activate          (Windows)   |   source .venv/bin/activate   (Mac/Linux)
       pip install -r requirements.txt
       streamlit run app.py
3. The browser opens at http://localhost:8501

## Passwords
Default app password: star123  |  admin password: adminstar
To change them: copy `.streamlit/secrets.toml.example` to `.streamlit/secrets.toml` and edit the two lines.

## Your data
Stored in `star_builders.db` (SQLite) next to app.py - back this file up.
First time: Sites & Excel tab > Import Excel > choose "Star - P&L Report.xlsx".

## Files
app.py   - screens (login, entries, add/edit, reports, sites & Excel)
core.py  - database, carry-forward of advances, reports, Excel import/export
