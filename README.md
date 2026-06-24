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
- Optional GitHub backup of JSON drafts to a private drafts repo
- GitHub recovery: search saved drafts by presenter name and load the selected draft back into the app
- An optional facilitator-notes appendix slide

The default content is prefilled with the OxyKids journal club example. Version 0.2.4 adds GitHub draft recovery so users can reaccess saved JSON files without manually downloading them from GitHub.

## Files

```text
journal_club_builder/
├── app.py                 # Streamlit app
├── pptx_builder.py        # PowerPoint generation functions
├── docx_builder.py        # One-page Word summary generation functions
├── feedback_config.py     # Fixed REDCap feedback URLs used in exports
├── github_storage.py      # Optional GitHub draft backup/recovery
├── slide_schema.py        # Slide fields, limits, defaults, and helper text
├── requirements.txt       # Python dependencies
├── README.md              # This file
├── .gitignore
└── .streamlit/
    └── config.toml        # Basic Streamlit theme
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

## Upload to GitHub

1. Create a new GitHub repository.
2. Upload all files in this folder to the repository.
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

## Optional GitHub draft backup and recovery

This feature lets the app save editable JSON drafts to a separate private GitHub repository, then reload them later from inside Streamlit.

Recommended setup:

1. Create a separate private GitHub repository named:

```text
journal-club-drafts
```

2. Create a fine-grained GitHub personal access token. Limit it to only the private `journal-club-drafts` repo. Give it:

```text
Contents: Read and write
```

3. Add these secrets in Streamlit Community Cloud under **App → Settings → Secrets**:

```toml
[github]
token = "github_pat_YOUR_TOKEN_HERE"
repo = "YOUR-GITHUB-USERNAME/journal-club-drafts"
branch = "main"
base_path = "drafts"
```

For local testing, copy `.streamlit/secrets.toml.example` to `.streamlit/secrets.toml` and fill in the same values. Do not commit a real `.streamlit/secrets.toml` file.

### Save a draft to GitHub

In the app, open **Backup draft to GitHub**, enter:

- Presenter name
- Session title

Then click **Save draft to GitHub**.

The app saves files like:

```text
drafts/2026-06-22_jane-smith_oxykids-trial.json
```

### Reload a draft from GitHub

In the app, open **Reload draft from GitHub**.

1. Enter the presenter name.
2. Click **Find saved drafts**.
3. Choose a draft from the list.
4. Click **Load selected draft**.

The app will refill all slide fields from the saved JSON draft.

Privacy note: the app searches by presenter name in the filename. For a small trusted education group, this is usually sufficient. For a larger or less trusted audience, consider adding a recovery code or email-based lookup later.

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
5. Use **Backup draft to GitHub** if they want a safety-net copy.
6. Download the PowerPoint.
7. Download the JSON draft if they want a local copy.
8. Use **Reload draft from GitHub** later if they need to recover prior work.


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


## Version 0.2.6 note

Slide 4 now uses a stable row-by-row table editor instead of `st.data_editor`. This avoids the confusing behavior where a table cell sometimes had to be deleted or typed twice before the change appeared to stick. Users can still edit the table columns, add rows, remove rows, and export the table to PowerPoint.


## Printable planning worksheet

Version 0.2.6 adds a pen-and-paper planning worksheet for residents who prefer to print the article and draft by hand before using the app.

In the app export panel, click:

```text
Download printable planning form
```

The worksheet mirrors the Streamlit fields and includes:

- slide-by-slide sections
- field labels matching the app
- required/optional status
- word limits, line limits, and row limits
- blank writing space
- a Slide 4 results table template

A standalone copy is also included in this repository:

```text
journal_club_printable_planning_form.docx
```

Residents can fill out the paper form, then transfer the final text into the app to generate the standardized PowerPoint and one-page summary.
