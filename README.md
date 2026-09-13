# Bunny Reading

A mobile-first pastel personal reading library built with Streamlit.

## Run locally

```bash
py -m pip install -r requirements.txt
py -m streamlit run app.py
```

## First-book content

The first demo book is stored in the `DEMO_BOOKS` constant near the top of `app.py`.

Later, replace the placeholder chapter `content` values with the final Product Development Interview Prep content, or move book data into JSON files under a future `data/` folder.

## Adding future books

The UI is not hard-coded around one book. Imported PDF, DOCX, TXT, pasted text, or webpage content is converted into the same reusable book/chapter data model.

For a larger next version, move:
- book data -> `data/books/*.json`
- extraction helpers -> `services/importers.py`
- page renderers -> `views/`
- persistence -> database/repository layer
