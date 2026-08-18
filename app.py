"""School ICT budgets in the Samagra Shiksha PAB minutes.

A companion to the NIPUN Bharat PAB app, built on the same workbook
discipline and the same visual language (ui.py is generated from that
app's dashboard.py by build_ict_ui.py, so the two cannot drift).

What this app publishes, and what it refuses to:

- It reads PAB_UDISE_ICT_master.xlsx and nothing else for money. Every
  figure on every page is a figure printed in a PAB minutes annexure,
  reached by a chain that closed against that document's own printed
  Sub Total and Total of ICT and Digital Initiatives line. Nothing is
  estimated, interpolated or back-filled.
- A year is either COMPLETE, meaning every state has been read and
  every block reconciles, or it is IN PROGRESS and is labelled as such
  on the page, everywhere it appears. A part-extracted year must never
  be able to read as a national total. That is why YEARS_COMPLETE is a
  separate constant from the years present in the workbook.
- A state absent from a year is not the same as a state with no data.
  Several states genuinely made no new school-ICT ask, and their
  minutes say so. NO_ASK records them with the evidence, so the app can
  say "asked for nothing" rather than showing a hole.
- Money and execution are different measurements taken from different
  tables a year apart, so they live in different tabs and are never
  silently added together.

Run locally with: streamlit run ict_dashboard/app.py
"""
import re
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st

from ui import (CARD, CSF_BLUE, CSF_YELLOW, GRID, INK, INK2, MONO, MUTED,
                PAPER, QUIET, SANS, SEQ, SERIES, STATUS, as_text, eyebrow,
                inject_css, right_align, section, table_csv)

ROOT = Path(__file__).parent
WB = ROOT / "PAB_UDISE_ICT_master.xlsx"
LINKS = ROOT / "PAB_document_links.xlsx"

# Years whose extraction is closed: every state read, every block
# reconciled against its own printed total. Only these are allowed to
# carry a national headline. Anything else in the workbook is shown,
# but always labelled with how many states it covers.
YEARS_COMPLETE = ["2024-25", "2025-26", "2026-27"]

# Blocks that are neither read nor confirmed absent, because the source
# we hold cannot answer. A third category, and it earns its own
# constant: rolling it into NO_ASK would claim a state asked for
# nothing when the truth is that we cannot see the page.
KNOWN_GAPS = {
    ("2024-25", "Mizoram", "Secondary"):
        "The minutes PDF contains only the ODD printed pages of its "
        "Budget Demand (PDF p25 is printed page 1, p26 is page 3, and "
        "so on). Printed page 38, the only place a Secondary ICT block "
        "could sit, was never scanned. No render resolution recovers "
        "it; it needs a re-fetch of the document. Mizoram's Elementary "
        "figure is complete, because that block happens to print "
        "entirely on one odd page.",
}

# States that appear in no costing row for a year because their minutes
# print no school-ICT ask at all, with what settles it. Distinguishing
# these from an extraction gap is the whole reason the app can quote a
# state count without hedging.
NO_ASK = {
    ("2026-27", "Chandigarh"): "State Plan summary prints ICT at 0.00",
    ("2026-27", "Delhi"): "State Plan summary prints ICT at 0.00",
    ("2026-27", "Goa"): "State Plan summary prints ICT at 0.00",
    ("2026-27", "Gujarat"): "State Plan summary prints ICT at 0.00",
    ("2026-27", "Lakshadweep"):
        "State Plan summary lists no ICT row (it lists only sub components "
        "carrying a figure)",
    ("2026-27", "Nagaland"):
        "State Plan summary lists no ICT row (it lists only sub components "
        "carrying a figure)",
    ("2026-27", "Puducherry"):
        "State Plan summary lists no ICT row (it lists only sub components "
        "carrying a figure)",
    ("2025-26", "West Bengal"):
        "Budget Demand is fully digital and prints no school ICT block; its "
        "ICT appears only in the spillover tables",
    ("2025-26", "Lakshadweep"):
        "Budget Demand is fully digital and prints no school ICT block; its "
        "ICT appears only in the spillover tables",
    ("2024-25", "Goa"):
        "No Total of ICT and Digital Initiatives among the boundary lines "
        "on any of its 20 Budget Demand pages; its only ICT is Technology "
        "Support to TEIs at p32, which is teacher education",
    ("2024-25", "Lakshadweep"):
        "No Total of ICT and Digital Initiatives among the boundary lines "
        "on any of its 15 Budget Demand pages; its only ICT is Technology "
        "Support to TEIs at p33, which is teacher education",
    ("2024-25", "Delhi"):
        "No Total of ICT and Digital Initiatives among the boundary lines "
        "on any of its 25 Budget Demand pages. Its p84 ICT is labs at the "
        "SCERT and 9 DIETs, and the full ICT table at p18-20 is a "
        "prior-year spillover annexure, not a costing sheet",
    ("2023-24", "Karnataka"):
        "All 93 Total of lines across its costing sheet were enumerated "
        "and none is ICT. Its only ICT is a BRC facility and DIET/SCERT "
        "technology support; the Total for ICT on p21 belongs to the "
        "prior-year Spill Over report, not a costing sheet",
    ("2023-24", "Delhi"):
        "All 30 Total of lines across its costing sheet were enumerated "
        "and none is ICT. Smart Classroom and Digital Hardware appear "
        "only on p17, in the March 2022-23 Spill Over table",
    ("2023-24", "Lakshadweep"):
        "Digital Hardware appears only in its p28-29 Spill Over report. "
        "Indicative rather than proven, because its costing sheet OCR is "
        "too damaged to enumerate the boundary lines",
}

COMPONENTS = ["ICT Lab", "Smart Classroom", "Digital Library",
              "Computer Devices", "Teacher Tablets", "Other ICT"]
COMP_COLOR = {
    "ICT Lab": CSF_BLUE, "Smart Classroom": CSF_YELLOW,
    "Digital Library": "#3b82f6", "Computer Devices": "#10b981",
    "Teacher Tablets": "#ec4899", "Other ICT": "#a1a1aa",
}
LEVELS = ["Elementary", "Secondary"]
NATURES = {"NR": "Non recurring", "R": "Recurring"}

# UDISE Plus measures, in the order a reader wants them rather than
# alphabetically. The slugs are udise_extract.py's; the labels are what
# goes on the page, and the raw printed header stays in each table's
# "As printed" column so the mapping can always be checked.
UD_LABEL = {
    "schools_with_computer": "Schools with a computer",
    "schools_with_functional_computer":
        "Schools with a working computer for teaching",
    "schools_with_internet": "Schools with internet",
    "schools_with_smart_classroom": "Schools with a smart classroom",
    "schools_with_digital_library": "Schools with a digital library",
    "govt_schools_with_ict_lab": "Government schools with an ICT lab",
    "govt_schools_with_functional_ict_lab":
        "Government schools with a working ICT lab",
    "govt_aided_schools_with_ict_lab":
        "Government aided schools with an ICT lab",
    "govt_aided_schools_with_functional_ict_lab":
        "Government aided schools with a working ICT lab",
    "schools_with_electricity": "Schools with electricity",
    "schools_with_functional_electricity":
        "Schools with working electricity",
    "schools_total": "Schools in total",
    "govt_schools_total": "Government schools in total",
    "govt_aided_schools_total": "Government aided schools in total",
}
UD_METRICS = list(UD_LABEL)

st.set_page_config(page_title="School ICT in PAB minutes",
                   layout="wide", page_icon="🖥️")
inject_css()


# --------------------------------------------------------------- data
@st.cache_data
def load(mtime=None):
    cost = pd.read_excel(WB, sheet_name="Data_Costing")
    exec_ = pd.read_excel(WB, sheet_name="Data_Execution")
    for df in (cost, exec_):
        df["state"] = df["state"].astype(str).str.strip()
    cost["nature_label"] = cost["nature"].map(NATURES).fillna("Not stated")
    cost["level"] = cost["level"].fillna("Not stated")
    # Rs lakh throughout the source; Cr is the unit anyone reads in
    for col, out in (("p_amt", "proposed_cr"), ("a_amt", "approved_cr")):
        cost[out] = cost[col] / 100.0
    for col in ("approved_fin", "completed_fin", "surrendered_fin",
                "cancelled_fin", "balance_fin"):
        exec_[col.replace("_fin", "_cr")] = exec_[col] / 100.0
    return cost, exec_


@st.cache_data
def load_links(mtime=None):
    """Portal and Drive URLs per source document, keyed by local file
    name so a Budget row can offer its own source in one click."""
    if not LINKS.exists():
        return pd.DataFrame(columns=["Local file", "Portal URL",
                                     "Drive link"])
    d = pd.read_excel(LINKS, sheet_name="PAB documents")
    d = d[d["Local file"].notna()]
    return d[["Local file", "Portal URL", "Drive link"]].drop_duplicates(
        subset=["Local file"])


@st.cache_data
def load_udise(mtime=None):
    """The school census, from the workbook's own Data_UDISE sheet.

    Read from the workbook rather than from udise_ict.csv on purpose:
    shipping both would put two copies of the same figures in the repo,
    and they would drift the first time one was refreshed alone. The
    workbook is the deliverable, so it is the single source here too.
    Returns None when the sheet is absent, and the tab says so rather
    than showing an estimate.
    """
    try:
        return pd.read_excel(WB, sheet_name="Data_UDISE")
    except ValueError:
        return None


def _mt(p):
    return p.stat().st_mtime if p.exists() else 0


if not WB.exists():
    st.error(f"Workbook not found at {WB}. Build it with "
             "build_ict_workbook.py and copy it next to this app.")
    st.stop()

COST, EXEC = load(_mt(WB))
DOCLINKS = load_links(_mt(LINKS))
UD = load_udise(_mt(WB))

YEARS_IN_WB = sorted(COST["year"].dropna().unique())
YEARS_PARTIAL = [y for y in YEARS_IN_WB if y not in YEARS_COMPLETE]
ALL_STATES = sorted(set(COST["state"]) | set(EXEC["state"]))
UD_HAVE = set() if UD is None else set(UD["metric"])


def source_url(local_file):
    row = DOCLINKS[DOCLINKS["Local file"] == local_file]
    if row.empty:
        return None
    for col in ("Portal URL", "Drive link"):
        v = row.iloc[0][col]
        if isinstance(v, str) and v.startswith("http"):
            return v
    return None


def cr(v, dp=0):
    return "no data" if pd.isna(v) else f"{v:,.{dp}f}"


def year_note(years):
    """One sentence naming the scope of whatever is being shown, so a
    part-extracted year can never sit unlabelled next to a closed one."""
    parts = []
    for y in years:
        n = COST[COST["year"] == y]["state"].nunique()
        n += sum(1 for (yy, _s) in NO_ASK if yy == y)
        if y in YEARS_COMPLETE:
            parts.append(f"{y} complete, {n} of 36 states read")
        else:
            parts.append(f"{y} IN PROGRESS, {n} of 36 states read so far")
    return ". ".join(parts) + "."


def state_totals(years, states=None, components=None):
    d = COST[COST["year"].isin(years)]
    if states:
        d = d[d["state"].isin(states)]
    if components:
        d = d[d["component"].isin(components)]
    return d


X_AXIS = alt.Axis(labelColor=MUTED, titleColor=QUIET, labelFont=SANS,
                  labelFontSize=11, domainColor=GRID, tickColor=GRID)
# axis tick values are digits that mean something, so they are monospace
Y_AXIS = alt.Axis(labelColor=MUTED, titleColor=QUIET, labelFont=MONO,
                  labelFontSize=11, gridColor=GRID, domainColor=GRID,
                  tickColor=GRID)


def bar(df, x, y, color=None, tooltip=None, height=300, sort=None,
        xtitle="", ytitle=""):
    """`x` and `y` accept either a field string or an already-built
    alt.X / alt.Y. Wrapping a built channel in alt.X() again produces a
    spec whose `field` is itself a channel object, which Vega-Lite
    rejects with a message that names `repeat` and points nowhere near
    the real cause."""
    enc = dict(
        x=x if isinstance(x, alt.X) else alt.X(x, title=xtitle, sort=sort,
                                               axis=X_AXIS),
        y=y if isinstance(y, alt.Y) else alt.Y(y, title=ytitle,
                                               axis=Y_AXIS),
    )
    if isinstance(x, alt.X):
        enc["x"] = x.copy()
        enc["x"]["axis"] = X_AXIS
        enc["x"]["title"] = xtitle
    if color is not None:
        enc["color"] = color
    if tooltip:
        enc["tooltip"] = tooltip
    return (alt.Chart(df).mark_bar(cornerRadius=2)
            .encode(**enc).properties(height=height)
            .configure_view(strokeWidth=0)
            .configure_legend(labelFont=SANS, labelColor=INK2,
                              titleFont=SANS, titleColor=QUIET,
                              labelFontSize=11, titleFontSize=10,
                              symbolType="square"))


# -------------------------------------------------------------- header
st.markdown('<div class="eyebrow">Samagra Shiksha AWP&amp;B minutes</div>',
            unsafe_allow_html=True)
st.title("School ICT in the PAB minutes")
st.caption(
    "What every state and union territory asked for, and what the Project "
    "Approval Board actually approved, for school computer labs, smart "
    "classrooms and the rest of the ICT and Digital Initiatives head. "
    "Every figure is read off a printed annexure page and reconciles "
    "against that page's own printed total. " + year_note(YEARS_IN_WB))

tab_story, tab_nat, tab_exp, tab_run, tab_ground, tab_qual = st.tabs(
    ["The Story", "National Picture", "Explore & Compare",
     "Approved vs Spent", "Schools on the Ground", "Data Quality"])


# ==================================================== 1. THE STORY ====
with tab_story:
    latest = YEARS_COMPLETE[-1]
    prev = YEARS_COMPLETE[0] if len(YEARS_COMPLETE) > 1 else None
    cur = state_totals([latest])
    prop, appr = cur["proposed_cr"].sum(), cur["approved_cr"].sum()

    with section("The headline", "navy"):
        st.markdown(
            f"In **{latest}**, states and union territories asked the "
            f"Project Approval Board for **Rs {cr(prop)} Cr** of school "
            f"ICT money. The board approved **Rs {cr(appr)} Cr** of it, "
            f"which is **{appr / prop * 100:.0f} paise in the rupee**.")
        c = st.columns(4)
        c[0].metric("Proposed", f"Rs {cr(prop)} Cr")
        c[1].metric("Approved", f"Rs {cr(appr)} Cr")
        c[2].metric("Approved share", f"{appr / prop * 100:.0f}%")
        c[3].metric("States asking",
                    f"{cur['state'].nunique()} of 36")
        st.caption(
            "Rs Cr, from the Rs lakh printed in each annexure. The states "
            "not asking are not missing data. Their minutes print no "
            "school ICT ask at all, and the Data Quality tab names each "
            "one with the line that settles it.")

    with section("The gap is the story, not the total", "gold"):
        st.markdown(
            "The interesting number here is not what was asked or what "
            "was approved. It is the distance between them, and how "
            "unevenly it falls. A state can have most of its ask met "
            "while its neighbour has almost none, and the printed "
            "Coordinator Remarks say why in plain language, usually one "
            "of three reasons. The school already has a functional lab "
            "recorded on PRABANDH. The school's enrolment falls under "
            "the hundred pupil floor the ministry applies. Or the "
            "recurring cost was asked for schools whose equipment is not "
            "yet working, and recurring money follows functional status.")
        g = (state_totals([latest]).groupby("state", as_index=False)
             [["proposed_cr", "approved_cr"]].sum())
        g["share"] = g["approved_cr"] / g["proposed_cr"] * 100
        g = g[g["proposed_cr"] > 0].sort_values("share")
        st.altair_chart(
            bar(g, alt.X("state:N", sort=g["state"].tolist()),
                "share:Q",
                color=alt.Color(
                    "share:Q", scale=alt.Scale(range=SEQ), legend=None),
                tooltip=[alt.Tooltip("state:N", title="State"),
                         alt.Tooltip("proposed_cr:Q", title="Proposed Cr",
                                     format=",.1f"),
                         alt.Tooltip("approved_cr:Q", title="Approved Cr",
                                     format=",.1f"),
                         alt.Tooltip("share:Q", title="Approved %",
                                     format=".0f")],
                height=320, ytitle="Approved as a share of the ask, %"),
            use_container_width=True)
        st.caption(f"{latest}. Sorted from the least met ask to the most.")

    with section("What the money is for", "blue"):
        comp = (state_totals(YEARS_COMPLETE)
                .groupby(["year", "component"], as_index=False)
                [["proposed_cr", "approved_cr"]].sum())
        st.markdown(
            "Two things dominate. An **ICT Lab** is a room of computers "
            "with its own recurring grant for the staff and consumables "
            "that keep it running. A **Smart Classroom** is a projector "
            "and a screen in an ordinary classroom, far cheaper per room "
            "and far cheaper to keep going. The split between them is a "
            "choice about whether digital learning happens in a special "
            "room or in the room children already sit in.")
        st.altair_chart(
            bar(comp, "year:N", "approved_cr:Q",
                color=alt.Color("component:N", title="Component",
                                scale=alt.Scale(
                                    domain=list(COMP_COLOR),
                                    range=list(COMP_COLOR.values()))),
                tooltip=[alt.Tooltip("year:N", title="Year"),
                         alt.Tooltip("component:N", title="Component"),
                         alt.Tooltip("approved_cr:Q", title="Approved Cr",
                                     format=",.1f")],
                height=280, ytitle="Approved, Rs Cr"),
            use_container_width=True)
        st.caption("Approved outlay by component. " +
                   year_note(YEARS_COMPLETE))

    with section("Building it once, and keeping it running", "emerald"):
        nat = (state_totals(YEARS_COMPLETE)
               .groupby(["year", "nature_label"], as_index=False)
               [["proposed_cr", "approved_cr"]].sum())
        st.markdown(
            "Every ICT line is either **non recurring**, the one time "
            "cost of buying and installing the equipment, or "
            "**recurring**, the yearly cost of keeping it working. The "
            "recurring side is the one that decides whether a lab is "
            "still teaching children in five years, and it is the side "
            "the board most often trims, because it will only fund "
            "recurring cost for equipment PRABANDH records as "
            "functional.")
        st.altair_chart(
            bar(nat, "year:N", "approved_cr:Q",
                color=alt.Color("nature_label:N", title="",
                                scale=alt.Scale(
                                    domain=["Non recurring", "Recurring"],
                                    range=[CSF_BLUE, CSF_YELLOW])),
                tooltip=[alt.Tooltip("year:N", title="Year"),
                         alt.Tooltip("nature_label:N", title=""),
                         alt.Tooltip("approved_cr:Q", title="Approved Cr",
                                     format=",.1f")],
                height=260, ytitle="Approved, Rs Cr"),
            use_container_width=True)

    with section("And whether the money was used", "amber"):
        ex = EXEC[EXEC["year_filled"].isin(EXEC["year_filled"].unique())]
        ay = sorted(ex["year_filled"].unique())
        if ay:
            y = ay[-1]
            e = ex[ex["year_filled"] == y]
            ap, co = e["approved_cr"].sum(), e["completed_cr"].sum()
            st.markdown(
                f"Approval is not spending. The ministry's own execution "
                f"report as at **{y}** shows **Rs {cr(co)} Cr** "
                f"completed against **Rs {cr(ap)} Cr** standing "
                f"approved, which is **{co / ap * 100:.0f}%**. The rest "
                f"sits as balance to be carried into the next year, or "
                f"was surrendered outright. The Approved vs Spent tab "
                f"has this per state.")
            c = st.columns(4)
            c[0].metric("Approved to date", f"Rs {cr(ap)} Cr")
            c[1].metric("Completed", f"Rs {cr(co)} Cr")
            c[2].metric("Surrendered",
                        f"Rs {cr(e['surrendered_cr'].sum())} Cr")
            c[3].metric("Balance carried",
                        f"Rs {cr(e['balance_cr'].sum())} Cr")
            st.caption(
                "Read off the Non Recurring Activities Report printed in "
                "the FOLLOWING year's minutes, which is where a year's "
                "execution is first published. Its approval column is "
                "CUMULATIVE, meaning everything still open from earlier "
                "years plus this year's new sanction, so it is larger "
                "than the year's own approved outlay and the two must "
                "not be compared. Budget and execution are different "
                "measurements from different tables and are never added "
                "together on this page.")


# ============================================= 2. NATIONAL PICTURE ====
with tab_nat:
    with section("Choose the year", "plain"):
        c = st.columns([2, 3])
        n_year = c[0].selectbox("Year", YEARS_IN_WB,
                                index=len(YEARS_IN_WB) - 1)
        n_comp = c[1].multiselect(
            "Components", COMPONENTS,
            default=[k for k in COMPONENTS
                     if k in set(COST["component"])],
            help="Every component the costing sheets carry for this year.")
        if n_year not in YEARS_COMPLETE:
            st.warning(
                f"{n_year} is still being extracted. Its figures cover "
                f"{COST[COST['year'] == n_year]['state'].nunique()} of 36 "
                "states and must not be read as a national total.")

    d = state_totals([n_year], components=n_comp)

    with section("The year in four figures", "navy"):
        c = st.columns(4)
        c[0].metric("Proposed", f"Rs {cr(d['proposed_cr'].sum())} Cr")
        c[1].metric("Approved", f"Rs {cr(d['approved_cr'].sum())} Cr")
        share = (d["approved_cr"].sum() / d["proposed_cr"].sum() * 100
                 if d["proposed_cr"].sum() else float("nan"))
        c[2].metric("Approved share",
                    "no data" if pd.isna(share) else f"{share:.0f}%")
        c[3].metric("Budget lines read", f"{len(d):,}")
        st.caption(year_note([n_year]))

    with section("Where the money goes, by state", "blue"):
        g = (d.groupby("state", as_index=False)
             [["proposed_cr", "approved_cr"]].sum()
             .sort_values("approved_cr", ascending=False))
        st.altair_chart(
            bar(g, alt.X("state:N", sort=g["state"].tolist()),
                "approved_cr:Q",
                color=alt.value(CSF_BLUE),
                tooltip=[alt.Tooltip("state:N", title="State"),
                         alt.Tooltip("proposed_cr:Q", title="Proposed Cr",
                                     format=",.2f"),
                         alt.Tooltip("approved_cr:Q", title="Approved Cr",
                                     format=",.2f")],
                height=340, ytitle="Approved, Rs Cr"),
            use_container_width=True)
        g["Approved share %"] = (g["approved_cr"] / g["proposed_cr"]
                                 * 100).round(0)
        show = g.rename(columns={"state": "State / UT",
                                 "proposed_cr": "Proposed Rs Cr",
                                 "approved_cr": "Approved Rs Cr"})
        st.dataframe(
            right_align(as_text(show,
                                ["Proposed Rs Cr", "Approved Rs Cr"],
                                ",.2f").style,
                        ["Proposed Rs Cr", "Approved Rs Cr",
                         "Approved share %"]),
            use_container_width=True, hide_index=True)
        table_csv(g, f"ict_by_state_{n_year}")

    with section("Elementary against secondary", "gold"):
        lv = (d.groupby(["level", "nature_label"], as_index=False)
              [["proposed_cr", "approved_cr"]].sum())
        st.markdown(
            "Elementary is schools up to the highest class VIII, "
            "secondary is up to class XII. The annexure prints the split "
            "as two separate blocks with their own subtotals, which is "
            "why they can be reported apart without any apportioning.")
        st.altair_chart(
            bar(lv, "level:N", "approved_cr:Q",
                color=alt.Color("nature_label:N", title="",
                                scale=alt.Scale(
                                    domain=["Non recurring", "Recurring"],
                                    range=[CSF_BLUE, CSF_YELLOW])),
                tooltip=[alt.Tooltip("level:N", title="Level"),
                         alt.Tooltip("nature_label:N", title=""),
                         alt.Tooltip("approved_cr:Q", title="Approved Cr",
                                     format=",.2f")],
                height=260, ytitle="Approved, Rs Cr"),
            use_container_width=True)

    with section("Component by component", "pink"):
        cp = (d.groupby("component", as_index=False)
              [["proposed_cr", "approved_cr"]].sum()
              .sort_values("approved_cr", ascending=False))
        cp["Approved share %"] = (cp["approved_cr"] / cp["proposed_cr"]
                                  * 100).round(0)
        show = cp.rename(columns={"component": "Component",
                                  "proposed_cr": "Proposed Rs Cr",
                                  "approved_cr": "Approved Rs Cr"})
        st.dataframe(
            right_align(as_text(show, ["Proposed Rs Cr", "Approved Rs Cr"],
                                ",.2f").style,
                        ["Proposed Rs Cr", "Approved Rs Cr",
                         "Approved share %"]),
            use_container_width=True, hide_index=True)
        table_csv(cp, f"ict_by_component_{n_year}")


# ============================================ 3. EXPLORE & COMPARE ====
with tab_exp:
    with section("Pick one state for its detail, or several to compare",
                 "plain"):
        picked = st.multiselect(
            "States and union territories", ALL_STATES,
            default=ALL_STATES[:1],
            help="One state shows every printed budget line with its "
                 "source page. Two or more compare them.")
        e_years = st.multiselect("Years", YEARS_IN_WB,
                                 default=list(YEARS_IN_WB))
    if not picked or not e_years:
        st.info("Pick at least one state and one year.")
    elif len(picked) == 1:
        s = picked[0]
        d = state_totals(e_years, states=[s])
        with section(f"{s}, in total", "navy"):
            if d.empty:
                reasons = [f"{y}, {NO_ASK[(y, s)]}"
                           for y in e_years if (y, s) in NO_ASK]
                if reasons:
                    st.info(f"{s} printed no school ICT ask in the years "
                            f"selected. " + ". ".join(reasons) + ".")
                else:
                    st.warning(
                        f"No ICT rows extracted for {s} in the years "
                        "selected. See the Data Quality tab for whether "
                        "that year is complete.")
            else:
                c = st.columns(3)
                c[0].metric("Proposed",
                            f"Rs {cr(d['proposed_cr'].sum(), 2)} Cr")
                c[1].metric("Approved",
                            f"Rs {cr(d['approved_cr'].sum(), 2)} Cr")
                c[2].metric(
                    "Approved share",
                    f"{d['approved_cr'].sum() / d['proposed_cr'].sum() * 100:.0f}%"
                    if d["proposed_cr"].sum() else "no data")
        if not d.empty:
            with section("Every printed line, and the page it is on",
                         "blue"):
                t = d[["year", "component", "level", "nature_label",
                       "activity", "p_phy", "p_amt", "a_phy", "a_amt",
                       "source_file", "pdf_page"]].copy()
                t["Source"] = t.apply(
                    lambda r: f"{r['source_file']} p.{int(r['pdf_page'])}"
                    if pd.notna(r["pdf_page"]) else r["source_file"],
                    axis=1)
                t = t.rename(columns={
                    "year": "Year", "component": "Component",
                    "level": "Level", "nature_label": "R / NR",
                    "activity": "Printed line",
                    "p_phy": "Proposed units", "p_amt": "Proposed Rs lakh",
                    "a_phy": "Approved units", "a_amt": "Approved Rs lakh"})
                t = t.drop(columns=["source_file", "pdf_page"])
                st.dataframe(
                    right_align(
                        as_text(t, ["Proposed Rs lakh", "Approved Rs lakh"],
                                ",.3f").pipe(
                            lambda x: as_text(x, ["Proposed units",
                                                  "Approved units"],
                                              ",.0f")).style,
                        ["Proposed units", "Proposed Rs lakh",
                         "Approved units", "Approved Rs lakh"]),
                    use_container_width=True, hide_index=True, height=430)
                table_csv(d, f"ict_{re.sub(r'[^A-Za-z0-9]+', '_', s)}")
                st.caption(
                    "Rs lakh here, as printed on the page, so a figure can "
                    "be checked against the annexure without arithmetic. "
                    "The Source column names the PDF and the page.")
            with section("The source documents", "plain"):
                for f in sorted(d["source_file"].dropna().unique()):
                    u = source_url(f)
                    st.markdown(f"- [{f}]({u})" if u else f"- {f}")
    else:
        d = state_totals(e_years, states=picked)
        with section("Side by side", "navy"):
            g = (d.groupby(["state", "year"], as_index=False)
                 [["proposed_cr", "approved_cr"]].sum())
            st.altair_chart(
                bar(g, "year:N", "approved_cr:Q",
                    color=alt.Color("state:N", title="State",
                                    scale=alt.Scale(range=SERIES)),
                    tooltip=[alt.Tooltip("state:N", title="State"),
                             alt.Tooltip("year:N", title="Year"),
                             alt.Tooltip("approved_cr:Q",
                                         title="Approved Cr",
                                         format=",.2f")],
                    height=320, ytitle="Approved, Rs Cr"),
                use_container_width=True)
        with section("Approved outlay by state and year", "blue"):
            p = (d.pivot_table(index="state", columns="year",
                               values="approved_cr", aggfunc="sum")
                 .reset_index().rename(columns={"state": "State / UT"}))
            ycols = [c for c in p.columns if c != "State / UT"]
            st.dataframe(
                right_align(as_text(p, ycols, ",.2f").style, ycols),
                use_container_width=True, hide_index=True)
            table_csv(p, "ict_comparison")
            st.caption("Rs Cr approved. " + year_note(e_years))
        with section("And the same for what was asked", "gold"):
            p2 = (d.pivot_table(index="state", columns="year",
                                values="proposed_cr", aggfunc="sum")
                  .reset_index().rename(columns={"state": "State / UT"}))
            ycols = [c for c in p2.columns if c != "State / UT"]
            st.dataframe(
                right_align(as_text(p2, ycols, ",.2f").style, ycols),
                use_container_width=True, hide_index=True)


# =========================================== 4. APPROVED VS SPENT ====
with tab_run:
    st.markdown(
        "Approval is a decision. Execution is what happened next. The "
        "ministry prints each year's execution in the **following** "
        "year's minutes, as a Non Recurring Activities Report, so this "
        "tab lags the budget tabs by one year and its figures come from "
        "a different table entirely.\n\n"
        "One thing to hold on to before reading any figure here. The "
        "report's approval column is **cumulative**. It carries "
        "everything still open from earlier years alongside the current "
        "year's new sanction, which is why the approved figure on this "
        "tab is larger than the approved outlay on the budget tabs. The "
        "two are not the same measurement and subtracting one from the "
        "other means nothing. What this tab is good for is the "
        "**proportions**, that is how much of the money standing "
        "approved has been turned into working equipment, how much was "
        "given back, and how much is still waiting.")
    ex_years = sorted(EXEC["year_filled"].dropna().unique())
    if not ex_years:
        st.info("No execution data loaded yet.")
    else:
        with section("Choose the year", "plain"):
            r_year = st.selectbox("Execution year", ex_years,
                                  index=len(ex_years) - 1,
                                  key="exec_year")
        e = EXEC[EXEC["year_filled"] == r_year]
        with section("What was approved, and what was completed", "amber"):
            c = st.columns(5)
            c[0].metric("Approved to date",
                        f"Rs {cr(e['approved_cr'].sum())} Cr")
            c[1].metric("Completed", f"Rs {cr(e['completed_cr'].sum())} Cr")
            c[2].metric("Surrendered",
                        f"Rs {cr(e['surrendered_cr'].sum())} Cr")
            c[3].metric("Balance carried",
                        f"Rs {cr(e['balance_cr'].sum())} Cr")
            done = (e["completed_cr"].sum() / e["approved_cr"].sum() * 100
                    if e["approved_cr"].sum() else float("nan"))
            c[4].metric("Completed share",
                        "no data" if pd.isna(done) else f"{done:.0f}%")
            st.caption(
                f"Position as at {r_year}, read off the report printed in "
                f"the next year's minutes. Covers {e['state'].nunique()} "
                "states and union territories. Approved to date is "
                "cumulative and is not this year's approved outlay.")
        with section("State by state", "blue"):
            g = (e.groupby("state", as_index=False)
                 [["approved_cr", "completed_cr", "surrendered_cr",
                   "balance_cr"]].sum())
            g["Completed %"] = (g["completed_cr"] / g["approved_cr"]
                                * 100).round(0)
            g = g.sort_values("approved_cr", ascending=False)
            m = g.melt(id_vars="state",
                       value_vars=["completed_cr", "surrendered_cr",
                                   "balance_cr"],
                       var_name="kind", value_name="cr")
            m["kind"] = m["kind"].map({"completed_cr": "Completed",
                                       "surrendered_cr": "Surrendered",
                                       "balance_cr": "Balance"})
            st.altair_chart(
                bar(m, alt.X("state:N", sort=g["state"].tolist()), "cr:Q",
                    color=alt.Color("kind:N", title="",
                                    scale=alt.Scale(
                                        domain=["Completed", "Surrendered",
                                                "Balance"],
                                        range=[STATUS["good"], STATUS["critical"],
                                               MUTED])),
                    tooltip=[alt.Tooltip("state:N", title="State"),
                             alt.Tooltip("kind:N", title=""),
                             alt.Tooltip("cr:Q", title="Rs Cr",
                                         format=",.2f")],
                    height=340, ytitle="Rs Cr"),
                use_container_width=True)
            show = g.rename(columns={
                "state": "State / UT",
                "approved_cr": "Approved to date Rs Cr",
                "completed_cr": "Completed Rs Cr",
                "surrendered_cr": "Surrendered Rs Cr",
                "balance_cr": "Balance Rs Cr"})
            num = ["Approved to date Rs Cr", "Completed Rs Cr",
                   "Surrendered Rs Cr", "Balance Rs Cr"]
            st.dataframe(
                right_align(as_text(show, num, ",.2f").style,
                            num + ["Completed %"]),
                use_container_width=True, hide_index=True)
            table_csv(g, f"ict_execution_{r_year}")
        with section("By component", "pink"):
            cg = (e.groupby("component", as_index=False)
                  [["approved_cr", "completed_cr", "surrendered_cr",
                    "balance_cr"]].sum()
                  .sort_values("approved_cr", ascending=False))
            num = ["approved_cr", "completed_cr", "surrendered_cr",
                   "balance_cr"]
            show = cg.rename(columns=dict(
                zip(["component"] + num,
                    ["Component", "Approved to date Rs Cr",
                     "Completed Rs Cr",
                     "Surrendered Rs Cr", "Balance Rs Cr"])))
            cols = [c for c in show.columns if c != "Component"]
            st.dataframe(
                right_align(as_text(show, cols, ",.2f").style, cols),
                use_container_width=True, hide_index=True)


# ======================================= 5. SCHOOLS ON THE GROUND ====
with tab_ground:
    st.markdown(
        "Money approved is one measure. Whether a school actually has a "
        "working computer is another, and **UDISE Plus**, the ministry's "
        "annual school census, counts it independently of any budget "
        "document. This tab sets the two beside each other and "
        "deliberately stops there. It never divides one by the other, "
        "because the census counts schools in a reference year while the "
        "board approves rupees in a financial year, and a rate built "
        "across them would be an invented number wearing a decimal "
        "point.")
    if UD is None or UD.empty:
        with section("Not published yet", "plain"):
            st.info(
                "The workbook carries no Data_UDISE sheet yet. This tab "
                "fills itself in as soon as build_ict_workbook.py is run "
                "with udise_ict.csv present, and shows nothing rather "
                "than an estimate in the meantime.")
    else:
        with section("Choose a measure", "plain"):
            c = st.columns([3, 2, 2])
            u_metric = c[0].selectbox(
                "Measure", [m for m in UD_METRICS if m in UD_HAVE],
                format_func=lambda m: UD_LABEL.get(m, m),
                help="Counted by the school census, independently of any "
                     "budget document.")
            avail_u = sorted(UD[UD["metric"] == u_metric]["unit"].unique())
            # share first, because a count chart just ranks states by how
            # many schools they have and says almost nothing about ICT
            u_unit = c[1].radio(
                "Shown as", avail_u, horizontal=True,
                index=avail_u.index("percent") if "percent" in avail_u else 0,
                format_func=lambda u: ("share of schools"
                                       if u == "percent" else "school count"),
                help="Percentages are printed by the census itself and are "
                     "never computed here.")
            u_years = sorted(UD[UD["metric"] == u_metric]["year"].unique())
            u_year = c[2].selectbox("Census year", u_years,
                                    index=len(u_years) - 1)

        unit_t = "% of schools" if u_unit == "percent" else "schools"
        u = UD[(UD["metric"] == u_metric) & (UD["unit"] == u_unit)
               & (UD["year"] == u_year)]
        scope = "; ".join(sorted(u["level"].unique()))
        india = u[u["state"] == "India"]
        g = u[u["state"] != "India"].sort_values("value", ascending=False)

        with section("Every state, as the census counts it", "emerald"):
            if g.empty:
                st.info("No rows for that combination.")
            else:
                st.altair_chart(
                    bar(g, alt.X("state:N", sort=g["state"].tolist()),
                        "value:Q", color=alt.value(CSF_BLUE),
                        tooltip=[alt.Tooltip("state:N", title="State"),
                                 alt.Tooltip("value:Q", title=unit_t,
                                             format=",.1f")],
                        height=340,
                        ytitle=f"{UD_LABEL.get(u_metric, u_metric)}, "
                               f"{unit_t}"),
                    use_container_width=True)
                note = (f"India prints {india['value'].iloc[0]:,.1f} on "
                        "the same page, and the 36 state and union "
                        "territory values sum to that row exactly, which "
                        "is how this extraction is checked. "
                        if not india.empty else "")
                st.caption(note + f"Scope as printed, {scope}.")
                tb = (g[["state", "value", "printed_label", "pdf_file",
                         "pdf_page"]]
                      .rename(columns={"state": "State / UT",
                                       "value": "Value",
                                       "printed_label": "As printed",
                                       "pdf_file": "Source",
                                       "pdf_page": "Page"}))
                st.dataframe(
                    right_align(as_text(tb, ["Value"], ",.1f").style,
                                ["Value"]),
                    use_container_width=True, hide_index=True, height=330)
                table_csv(u, f"udise_{u_metric}_{u_year}")

        with section("The same measure across the census years", "blue"):
            tr = UD[(UD["metric"] == u_metric) & (UD["unit"] == u_unit)]
            pick = st.multiselect(
                "States to trace", sorted(set(tr["state"]) - {"India"}),
                default=list(g["state"].head(5)), key="ud_trace")
            t = tr[tr["state"].isin(pick + ["India"])]
            if not t.empty:
                st.altair_chart(
                    alt.Chart(t).mark_line(point=True, strokeWidth=2)
                    .encode(
                        x=alt.X("year:N", title="", axis=X_AXIS),
                        y=alt.Y("value:Q", title=unit_t, axis=Y_AXIS),
                        color=alt.Color("state:N", title="State",
                                        scale=alt.Scale(range=SERIES)),
                        tooltip=[alt.Tooltip("state:N", title="State"),
                                 alt.Tooltip("year:N", title="Year"),
                                 alt.Tooltip("value:Q", title=unit_t,
                                             format=",.1f")])
                    .properties(height=300)
                    .configure_view(strokeWidth=0)
                    .configure_legend(labelFont=SANS, labelColor=INK2,
                                      titleFont=SANS, titleColor=QUIET,
                                      labelFontSize=11, titleFontSize=10,
                                      symbolType="stroke"),
                    use_container_width=True)
            if u_metric.startswith("govt"):
                st.warning(
                    "The ICT lab table changes what it covers between "
                    "volumes. It reads Upper Primary, Secondary and "
                    "Higher Secondary sections in 2021-22 and again in "
                    "2025-26, but Middle and Secondary sections in "
                    "between, which moves the denominator underneath the "
                    "figure. Read this as three separate stretches, not "
                    "one trend.")

        with section("The money and the count, side by side", "gold"):
            st.markdown(
                "Two measurements of the same states, drawn from two "
                "unrelated documents. A state high on one axis and low "
                "on the other is worth a question rather than a "
                "conclusion, because one year of approvals is small "
                "against a stock of schools built up over many.")
            c = st.columns(2)
            b_year = c[0].selectbox("Budget year", YEARS_IN_WB,
                                    index=len(YEARS_IN_WB) - 1,
                                    key="ud_budget_year")
            m_year = c[1].selectbox("Census year", u_years,
                                    index=len(u_years) - 1,
                                    key="ud_census_year")
            money = (state_totals([b_year]).groupby("state", as_index=False)
                     ["approved_cr"].sum())
            cen = (UD[(UD["metric"] == u_metric) & (UD["unit"] == u_unit)
                      & (UD["year"] == m_year) & (UD["state"] != "India")]
                   [["state", "value"]])
            both = money.merge(cen, on="state", how="inner")
            if both.empty:
                st.info("No states appear in both of those selections.")
            else:
                st.altair_chart(
                    alt.Chart(both).mark_circle(size=120, opacity=0.75)
                    .encode(
                        x=alt.X("value:Q", axis=Y_AXIS,
                                title=f"{UD_LABEL.get(u_metric, u_metric)}"
                                      f", census {m_year} ({unit_t})"),
                        y=alt.Y("approved_cr:Q", axis=Y_AXIS,
                                title=f"ICT approved Rs Cr, {b_year}"),
                        color=alt.value(CSF_BLUE),
                        tooltip=[alt.Tooltip("state:N", title="State"),
                                 alt.Tooltip("value:Q", title="Census",
                                             format=",.1f"),
                                 alt.Tooltip("approved_cr:Q",
                                             title="Approved Cr",
                                             format=",.2f")])
                    .properties(height=340).configure_view(strokeWidth=0),
                    use_container_width=True)
                st.caption(
                    f"{len(both)} states appear in both. Horizontal is "
                    "the census, vertical is the board. No line is fitted "
                    "through these points and no ratio is taken across "
                    "the two axes.")

        with section("What the census prints, and what it does not",
                     "amber"):
            st.markdown(
                "- Counts and percentages are reproduced exactly as the "
                "report prints them. A percentage is never turned into a "
                "count, because the denominator the report used is not "
                "always the one on the same page.\n"
                "- Every figure was checked twice over. The 36 state "
                "values sum to the printed India row for every count "
                "measure in every year, and every printed percentage "
                "agrees with its own printed count and denominator.\n"
                "- Table numbering moves between volumes. The computers "
                "and digital section is Section 10 in the 2021-22 report "
                "and Section 9 from 2022-23 onward, so the As printed "
                "column names the table each figure actually came from "
                "rather than assuming a fixed number.\n"
                "- One source defect is carried rather than hidden. The "
                "2025-26 smart classroom table prints a damaged "
                "column-number band on the page itself, so those rows "
                "were read by column position and reconciled, and their "
                "note says so.")


# ============================================== 6. DATA QUALITY ====
with tab_qual:
    st.markdown(
        "Everything on the other tabs rests on this one. A figure is "
        "publishable here only if it was read off a printed page and its "
        "block closed against that page's own printed Sub Total and "
        "Total of ICT and Digital Initiatives line. Nothing is inferred "
        "from arithmetic alone.")

    with section("Coverage, year by year", "navy"):
        rows = []
        for y in YEARS_IN_WB:
            d = COST[COST["year"] == y]
            noask = sum(1 for (yy, _s) in NO_ASK if yy == y)
            rows.append({
                "Year": y,
                "Status": ("Complete" if y in YEARS_COMPLETE
                           else "In progress"),
                "States with an ICT ask": d["state"].nunique(),
                "States printing no ask": noask,
                "States accounted for": d["state"].nunique() + noask,
                "Budget lines": len(d),
                "Approved Rs Cr": round(d["approved_cr"].sum(), 2)})
        cov = pd.DataFrame(rows)
        st.dataframe(
            right_align(as_text(cov, ["Approved Rs Cr"], ",.2f").style,
                        ["States with an ICT ask", "States printing no ask",
                         "States accounted for", "Budget lines",
                         "Approved Rs Cr"]),
            use_container_width=True, hide_index=True)
        st.caption(
            "36 states and union territories submit a plan each year. A "
            "year is called complete only when every one of them is "
            "either read or confirmed to print no ask.")

    with section("The states that asked for nothing", "gold"):
        st.markdown(
            "These are not gaps. Each one was opened and checked, and "
            "the line that settles it is given here so it can be "
            "rechecked.")
        na = pd.DataFrame(
            [{"Year": y, "State / UT": s, "What settles it": why}
             for (y, s), why in sorted(NO_ASK.items())])
        st.dataframe(na, use_container_width=True, hide_index=True)

    with section("The one thing the sources cannot answer", "pink"):
        st.markdown(
            "A year marked complete still has to admit what it could "
            "not see. This is not a state that asked for nothing and "
            "not a block still to be read. It is a page that was never "
            "scanned into the document the ministry published.")
        kg = pd.DataFrame(
            [{"Year": y, "State / UT": s, "Level": lv, "Why": why}
             for (y, s, lv), why in sorted(KNOWN_GAPS.items())])
        st.dataframe(kg, use_container_width=True, hide_index=True)
        st.caption(
            "The affected state's other level is unaffected and is "
            "published normally. No figure has been estimated to fill "
            "the gap.")

    with section("Where each figure came from", "blue"):
        src = (COST.groupby(["year", "source_file"], as_index=False)
               .agg(lines=("activity", "size"),
                    states=("state", "nunique"),
                    approved_cr=("approved_cr", "sum")))
        src["Link"] = src["source_file"].map(source_url)
        src = src.sort_values(["year", "source_file"])
        show = src.rename(columns={
            "year": "Year", "source_file": "Source document",
            "lines": "Budget lines", "states": "States",
            "approved_cr": "Approved Rs Cr"})
        st.dataframe(
            right_align(as_text(show, ["Approved Rs Cr"], ",.2f").style,
                        ["Budget lines", "States", "Approved Rs Cr"]),
            use_container_width=True, hide_index=True, height=380,
            column_config={"Link": st.column_config.LinkColumn(
                "Source", display_text="open")})
        table_csv(src, "ict_sources")
        st.caption(
            "Every budget line in this app carries the document it was "
            "read from and the page number within it. Where the ministry "
            "portal or the shared drive has a copy, the link opens it.")

    with section("How to read a figure back to its page", "emerald"):
        st.markdown(
            "1. Find the state and year in **Explore & Compare** and "
            "select the single state. Its table lists every printed "
            "line with a Source column naming the PDF and page.\n"
            "2. Open that document from the list beneath the table.\n"
            "3. Turn to the page. The line will be there, in Rs lakh, "
            "with the same units, and the block it sits in will total to "
            "the same figure this app shows.\n\n"
            "If any of those three steps fails, that is a defect worth "
            "reporting, not a rounding difference.")

    with section("What this app does not claim", "amber"):
        st.markdown(
            "- It does not estimate. A state with no printed figure "
            "shows as having none.\n"
            "- It does not merge budget with execution, or either with "
            "the school census. Those are three separate measurements "
            "and each has its own tab.\n"
            "- It does not adjust for inflation or population, so a "
            "figure is the rupee amount the board printed, nothing "
            "more.\n"
            "- Years still being extracted are labelled in every place "
            "they appear, and are excluded from any national headline.")

    st.caption(
        f"Built from {WB.name}. Costing rows {len(COST):,}, execution "
        f"rows {len(EXEC):,}. Complete years {', '.join(YEARS_COMPLETE)}."
        + (f" In progress {', '.join(YEARS_PARTIAL)}."
           if YEARS_PARTIAL else ""))
