# School ICT in the PAB minutes

A Streamlit app publishing what every Indian state and union territory
asked for, and what the Samagra Shiksha Project Approval Board actually
approved, under the **ICT and Digital Initiatives** head: school
computer labs, smart classrooms, digital libraries, teacher tablets and
the recurring grants that keep them running.

Live at https://pab-ict.streamlit.app

Companion to the [NIPUN Bharat PAB app](https://nipun-pab.streamlit.app),
built on the same workbook discipline and the same visual language.

## The standard

Fidelity to print. Every figure in this app is a figure printed in a PAB
minutes annexure, reached by a chain that closed against that page's own
printed `Sub Total` and `Total of ICT and Digital Initiatives` line.
Nothing is estimated, interpolated or back-filled, and every row carries
the source document and page number so it can be read back.

Three consequences worth knowing before quoting anything:

- **A year is complete or it is labelled.** Only years in
  `YEARS_COMPLETE` carry a national headline. Any year still being
  extracted is marked in every place it appears and excluded from
  totals, so a part-extracted year cannot read as a national figure.
- **A state with no rows is not a state with no data.** Several states
  print no school ICT ask at all. `NO_ASK` records each one with the
  line that settles it, and the app says "asked for nothing" rather
  than showing a hole.
- **Budget and execution are never added.** The execution figures come
  from the Non Recurring Activities Report printed in the *following*
  year's minutes, and their approval column is cumulative, carrying
  everything still open from earlier years. It is a larger number than
  the year's approved outlay and means something different.

## Files

| File | What it is |
|---|---|
| `app.py` | The app. Six tabs. |
| `ui.py` | The visual language. **Generated** from the NIPUN app's `dashboard.py` by `build_ict_ui.py` in the parent project. Do not edit here. |
| `PAB_UDISE_ICT_master.xlsx` | The data. Built by `build_ict_workbook.py`. |
| `PAB_document_links.xlsx` | Portal and Drive URLs per source document. |
| `udise_ict.csv` | UDISE Plus school infrastructure counts. Optional; the Schools on the Ground tab degrades gracefully without it. |

## Running locally

```
pip install -r requirements.txt
streamlit run app.py
```

## Updating

Rebuild the workbook in the parent project, copy it here, commit, push.
Streamlit Cloud redeploys on push.
