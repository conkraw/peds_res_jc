# Journal Club PowerPoint Builder

A Streamlit app for building standardized resident journal club PowerPoints and one-page session summaries.

The app provides:

- Simple sidebar slide navigation with the editable fields in the main workspace
- Word, line, and table-length limits
- A standardized PowerPoint export
- A one-page Word summary export with the fixed feedback website link and QR code
- A final feedback slide with a fixed REDCap website link and QR code
- A Slide 4 visual option: results table, big-number card, simple bar chart, or no visual
- Download/reload of editable JSON drafts
- Optional backup of editable JSON drafts to a private GitHub repo
- An optional facilitator-notes appendix slide

The default content is prefilled with the OxyKids journal club example. Version 0.2.3 adds optional GitHub draft backup while keeping the resident-facing app simple.

## Files

```text
journal_club_builder/
├── app.py                         # Streamlit app
├── pptx_builder.py                # PowerPoint generation functions
├── docx_builder.py                # One-page Word summary generation functions
├── github_storage.py              # Optional GitHub JSON draft backup
├── feedback_config.py             # Fixed REDCap feedback URLs used in exports
├── slide_schema.py                # Slide fields, limits, defaults, and helper text
├── requirements.txt               # Python dependencies
├── README.md                      # This file
├── .gitignore
└── .streamlit/
    ├── config.toml                # Basic Streamlit theme
    └── secrets.toml.example       # Example only; do not put a real token in GitHub
```

## Run locally

From the project folder:

```bash
python -m venv .venv
source .venv/bin/activate   # Mac/Linux
# .venv\Scripts\activate    # Windows PowerShell

pip install -r requirements.txt
streamlit run app.py
```

## Upload the app to GitHub

1. Create a GitHub repository for the Streamlit app, for example `journal-club-builder`.
2. Upload all files in this folder to that repository.
3. Commit the files to the `main` branch.

## Deploy on Streamlit Community Cloud

1. Go to Streamlit Community Cloud.
2. Choose **New app**.
3. Select your GitHub repository.
4. Set the main file path to:

```text
app.py
```

5. Deploy.

## Optional: save resident JSON drafts to GitHub

This feature lets the app save a backup JSON file into a GitHub repo when the user clicks **Save draft to GitHub**.

Recommended setup:

```text
App repo:      journal-club-builder
Drafts repo:   journal-club-drafts   # private
```

Use a separate **private** drafts repo because the JSON may contain presenter names, session titles, and internal educational content.

### Step 1 — Create the private drafts repo

1. In GitHub, create a new repository.
2. Name it something like:

```text
journal-club-drafts
```

3. Set visibility to **Private**.
4. Add a README if you want.
5. Keep the default branch as `main`.

The app will automatically create a `drafts/` folder the first time it saves a draft.

### Step 2 — Create a GitHub token

Use a fine-grained personal access token if possible.

Recommended token settings:

```text
Repository access:
Only selected repositories → journal-club-drafts

Repository permissions:
Contents: Read and write
Metadata: Read-only
```

Copy the token right away. You will not be able to see it again after leaving the token page.

### Step 3 — Add secrets to Streamlit Community Cloud

In your deployed Streamlit app settings, add these secrets:

```toml
[github]
token = "github_pat_YOUR_TOKEN_HERE"
repo = "YOUR-GITHUB-USERNAME/journal-club-drafts"
branch = "main"
base_path = "drafts"
```

Do not put the real token in the GitHub repo.

### Step 4 — Local testing secrets, optional

For local testing only, copy:

```text
.streamlit/secrets.toml.example
```

to:

```text
.streamlit/secrets.toml
```

Then put your real token in `secrets.toml`.

The `.gitignore` file already prevents `.streamlit/secrets.toml` from being committed.

### Step 5 — How saved files are named

When someone clicks **Save draft to GitHub**, the app saves a file like:

```text
drafts/2026-06-22_jane-smith_oxykids-trial.json
```

The filename uses:

```text
today's date + presenter name + session title
```

If the same presenter saves the same session again on the same date, the app updates the existing file instead of creating duplicates.

### Step 6 — Restoring a lost draft

1. Open the private `journal-club-drafts` repo.
2. Go to the `drafts/` folder.
3. Find the JSON file by date, presenter, or session title.
4. Download the JSON file.
5. In the Streamlit app, open **Advanced: drafts/reset**.
6. Upload the JSON file using **Load a saved draft JSON**.

The app can load both older plain deck JSON files and newer GitHub backup JSON files with metadata.

## How to customize slides

Edit `slide_schema.py` to change:

- Slide names
- Field labels
- Default text
- Word limits
- Line limits
- Helper text
- Slide 4 visual fields

Each slide has an `id`, a `label`, and a list of `fields`.

Example field:

```python
{
    "key": "main_result",
    "label": "Main result headline",
    "type": "text",
    "required": True,
    "max_words": 18,
    "default": "The 88% group met discharge criteria sooner.",
}
```

## How to customize PowerPoint layout

Edit `pptx_builder.py` to change:

- Fonts
- Colors
- Layouts
- Slide sizing
- Table formatting
- Footer text
- Visual style for Slide 4
- Final feedback slide layout and QR code placement

The app currently builds a clean widescreen PowerPoint using `python-pptx`.

## Notes about speaker notes

The app exports facilitator notes as an editable appendix slide rather than hidden PowerPoint speaker notes. This keeps the export reliable across `python-pptx` versions and makes notes easy to review.

## Recommended workflow for residents

1. Open the app.
2. Choose a slide from the simple sidebar list.
3. Complete the fields in the main workspace and keep each field within the displayed limits.
4. Use the Slide 4 results table unless there is a strong reason to use a different visual.
5. Optional: open **Backup draft to GitHub**, enter presenter name, and click **Save draft to GitHub**.
6. Download the PowerPoint.
7. Download the one-page Word summary.

## Feedback slide

The feedback slide is added automatically at the end of the PowerPoint. Residents do not edit it in Streamlit.

The PowerPoint and Word summary always use:

- QR code URL: `https://redcap.ctsi.psu.edu/surveys/?s=T9P4FPRYMJ3XL478`
- Display website link: `https://redcap.link/peds_res_jc_feedback`

## One-page Word summary

The **Download 1-page summary** button creates a compact `.docx` file with:

- Session title and article topic
- Teaching purpose
- PICO / study question
- Main result
- Clinical bottom line
- Trust/caution points
- Discussion questions
- Resident take-home sentence
