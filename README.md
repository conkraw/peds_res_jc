# Journal Club PowerPoint Builder

A Streamlit app for building standardized resident journal club PowerPoints.

The app provides:

- Slide-by-slide sidebar editing
- Word, line, and table-length limits
- A standardized PowerPoint export
- A Slide 4 visual option: results table, big-number card, simple bar chart, or no visual
- Download/reload of editable JSON drafts
- An optional facilitator-notes appendix slide

The default content is prefilled with the OxyKids journal club example.

## Files

```text
journal_club_builder/
├── app.py                 # Streamlit app
├── pptx_builder.py        # PowerPoint generation functions
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

The app currently builds a clean widescreen PowerPoint using `python-pptx`.

## Notes about speaker notes

The app exports facilitator notes as an editable appendix slide rather than hidden PowerPoint speaker notes. This keeps the export reliable across `python-pptx` versions and makes notes easy to review.

## Recommended workflow for residents

1. Open the app.
2. Move through the sidebar slide list from top to bottom.
3. Keep each field within the displayed limits.
4. Use the Slide 4 results table unless there is a strong reason to use a different visual.
5. Download the PowerPoint.
6. Download the JSON draft if they want to save their work and return later.
