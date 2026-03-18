
import streamlit as st
import pandas as pd
from datetime import timedelta
from io import BytesIO

# Optional CP-SAT (ortools)
try:
    from ortools.sat.python import cp_model
    ORTOOLS_OK = True
except Exception:
    ORTOOLS_OK = False

st.set_page_config(page_title="Coatinc De Meern - Planning Optimizer", layout="wide")
st.title("Coatinc De Meern - Planning Optimizer")

# --- Branding ---
import os as _os
_logo_path = _os.path.join(_os.path.dirname(__file__), "assets", "coatinc_logo.png")
if _os.path.exists(_logo_path):
    st.sidebar.image("assets/coatinc_logo.png", width=280)


st.markdown(
    """<style>
    /* Best-effort wrapping in Streamlit dataframes */
    div[data-testid='stDataFrame'] div[role='gridcell']{
        white-space: normal !important;
        line-height: 1.2em;
    }
    </style>""",
    unsafe_allow_html=True
)


# ---------- Sidebar controls ----------
st.sidebar.header("Basis instellingen")
max_m2 = st.sidebar.number_input("Max m² per dag", value=2000, step=100, min_value=1)
max_kleuren = st.sidebar.number_input("Max kleuren per dag", value=18, step=1, min_value=1)

st.sidebar.subheader("Leverdatum buffer (plandatum vóór leverdatum)")
min_dagen_voor_lever = st.sidebar.number_input(
    "Minimaal # dagen vóór leverdatum (hard)", value=2, step=1, min_value=2,
    help="Hard: productie moet uiterlijk LeverDatum - dit aantal dagen gepland staan."
)
pref_dagen_voor_lever = st.sidebar.number_input(
    "Voorkeur # dagen vóór leverdatum (soft)", value=2, step=1, min_value=1,
    help="Soft: planner-afspraak (bijv. 2 of 3 dagen). CP-SAT probeert hier zoveel mogelijk aan te voldoen."
)

st.sidebar.subheader("Fixeren")
fixeer_morgen = st.sidebar.checkbox("Fixeer ook morgen", value=True, help="Vandaag en eerder zijn altijd gefixeerd. Met dit vinkje wordt morgen ook op slot gezet.")

st.sidebar.subheader("Kalender")
exclude_weekends = st.sidebar.checkbox("Weekenden uitsluiten", value=True)
exclude_holidays = st.sidebar.checkbox("Feestdagen uitsluiten", value=True)
holiday_text = st.sidebar.text_area(
    "Feestdagen (YYYY-MM-DD, één per regel)",
    value="",
    help="Voorbeeld:\n2026-04-27\n2026-12-25"
)

st.sidebar.subheader("Ontvangstconstraint (alleen als Binnen = onwaar)")
use_ontvangst = st.sidebar.checkbox("Gebruik aanleverdatum regel", value=True)
min_dagen_ontvangst = st.sidebar.number_input("Min. dagen tussen aanleverdatum en plandatum (hard)", value=1, step=1, min_value=0)
pref_dagen_ontvangst = st.sidebar.number_input("Voorkeur dagen tussen aanleverdatum en plandatum (soft)", value=2, step=1, min_value=0)

st.sidebar.subheader("Optimizer")
optimizer = st.sidebar.selectbox("Methode", ["Heuristiek (snel)", "CP-SAT (beste oplossing)"])
time_limit = st.sidebar.slider("CP-SAT tijdslimiet (sec)", 5, 120, 20, step=5, disabled=(optimizer!="CP-SAT (beste oplossing)" or not ORTOOLS_OK))

uploaded = st.file_uploader("Upload planning Excel", type=["xlsx"])

# ---------- Helpers ----------
def parse_holidays(txt: str):
    dates = set()
    for line in txt.splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            dates.add(pd.to_datetime(line).normalize())
        except Exception:
            pass
    return dates

def fmt_m2(x):
    try:
        return f"{int(round(float(x))):,}".replace(",", ".")
    except Exception:
        return x

def fmt_df_m2(df, cols):
    for c in cols:
        if c in df.columns:
            df[c] = df[c].apply(fmt_m2)
    return df


def normalize_dates(df_in: pd.DataFrame) -> pd.DataFrame:
    """Strip times from all date/datetime columns (normalize to midnight)."""
    for c in df_in.columns:
        low = str(c).lower()
        if ("datum" in low) or low.endswith("dag"):
            if pd.api.types.is_datetime64_any_dtype(df_in[c]):
                df_in[c] = pd.to_datetime(df_in[c], errors="coerce").dt.normalize()
    return df_in

def format_dates_ddmmyyyy(df_in: pd.DataFrame) -> pd.DataFrame:
    """Format all date/datetime columns as dd-mm-jjjj strings for UI tables."""
    df2 = df_in.copy()
    for c in df2.columns:
        low = str(c).lower()
        if ("datum" in low) or low.endswith("dag"):
            if pd.api.types.is_datetime64_any_dtype(df2[c]):
                df2[c] = pd.to_datetime(df2[c], errors="coerce").dt.strftime("%d-%m-%Y")
    return df2




def make_column_config(df: pd.DataFrame):
    cfg = {}
    for col in df.columns:
        low = str(col).lower()
        # Long text: keep medium so sheet stays compact
        if low in ["omschrijving", "referentieklant"]:
            cfg[col] = st.column_config.TextColumn(col, width="medium")
        elif low in ["naam"]:
            cfg[col] = st.column_config.TextColumn(col, width="large")
        elif low in ["kleur", "orderid", "deelorderid", "volgnr", "prio", "ordertypeafkorting"]:
            cfg[col] = st.column_config.TextColumn(col, width="small")
        elif ("datum" in low) or low.endswith("dag"):
            # Dates are shown as dd-mm-jjjj strings (we pre-format), so treat as text
            cfg[col] = st.column_config.TextColumn(col, width="medium")
        else:
            # Default: large so most columns are fully visible
            cfg[col] = st.column_config.TextColumn(col, width="large")
    return cfg



def is_weekend(d):
    return d.weekday() >= 5

def make_working_days(start, end, excl_weekends, holidays_set, excl_holidays):
    days = pd.date_range(start=start, end=end, freq="D")
    out = []
    for d in days:
        dn = d.normalize()
        if excl_weekends and is_weekend(dn):
            continue
        if excl_holidays and dn in holidays_set:
            continue
        out.append(dn)
    return out

def safe_bool(x):
    # Excel can contain True/False, 1/0, 'WAAR'/'ONWAAR'
    if isinstance(x, bool):
        return x
    if pd.isna(x):
        return False
    s = str(x).strip().lower()
    if s in ("true","waar","1","yes","y","ja"):
        return True
    if s in ("false","onwaar","0","no","n","nee"):
        return False
    return False

# ---------- Main ----------
if not uploaded:
    st.info("Upload een Excel met minimaal kolommen: PlanDatum, LeverDatum, Kleur, M2, Binnen (en bij voorkeur AanleverDatum).")
    st.stop()

df = pd.read_excel(uploaded)

required = ["PlanDatum", "LeverDatum", "Kleur", "M2", "Binnen"]
missing = [c for c in required if c not in df.columns]
if missing:
    st.error(f"Kolommen ontbreken: {', '.join(missing)}")
    st.stop()

if "OrderID" not in df.columns:
    df["OrderID"] = range(1, len(df)+1)

# Parse dates
df["PlanDatum"] = pd.to_datetime(df["PlanDatum"], errors="coerce")
df["LeverDatum"] = pd.to_datetime(df["LeverDatum"], errors="coerce")
if "AanleverDatum" in df.columns:
    df["AanleverDatum"] = pd.to_datetime(df["AanleverDatum"], errors="coerce")

# Normalize Binnen
df["_BinnenBool"] = df["Binnen"].apply(safe_bool)

vandaag = pd.Timestamp.today().normalize()
morgen = vandaag + timedelta(days=1)

# Fix rules
df["Gefixeerd"] = df["PlanDatum"].dt.normalize() <= vandaag
if fixeer_morgen:
    df.loc[df["PlanDatum"].dt.normalize() == morgen, "Gefixeerd"] = True

# Locked days set (days that are fixed horizon = on slot)
locked_days = set(df.loc[df["Gefixeerd"], "PlanDatum"].dt.normalize().dropna().unique())

holidays = parse_holidays(holiday_text)

# Planning horizon: from tomorrow to max(leverdatum) + some slack
max_lever = df["LeverDatum"].max()
if pd.isna(max_lever):
    st.error("LeverDatum bevat lege/ongeldige waarden.")
    st.stop()

horizon_end = (max_lever + timedelta(days=14)).normalize()
start = (vandaag + timedelta(days=1)).normalize()

working_days = make_working_days(start, horizon_end, exclude_weekends, holidays, exclude_holidays)
# Remove locked days from candidate planning days (no new orders may be placed there)
candidate_days = [d for d in working_days if d not in locked_days]

if len(candidate_days) == 0:
    st.error("Geen planbare dagen in de horizon (mogelijk alles geblokkeerd door weekends/feestdagen/fixaties).")
    st.stop()

# Precompute fixed day usage
fixed = df[df["Gefixeerd"]].copy()
fixed["PlanDatumN"] = fixed["PlanDatum"].dt.normalize()
fixed_m2 = fixed.groupby("PlanDatumN")["M2"].sum().to_dict()
fixed_colors = fixed.groupby("PlanDatumN")["Kleur"].nunique().to_dict()

# Day warning for fixed overload
overload_days = []
for d in sorted(locked_days):
    m2d = float(fixed_m2.get(d, 0.0))
    cd = int(fixed_colors.get(d, 0))
    if m2d > max_m2 or cd > max_kleuren:
        overload_days.append((d, m2d, cd))

if overload_days:
    st.warning("⚠️ Dagmelding: gefixeerde orders overschrijden daglimieten (m² en/of kleuren).")
    warn_df = pd.DataFrame(overload_days, columns=["Dag", "m2_gefixeerd", "kleuren_gefixeerd"])
    warn_df["m2_gefixeerd"] = warn_df["m2_gefixeerd"].apply(fmt_m2)
    st.dataframe(format_dates_ddmmyyyy(warn_df), use_container_width=True, hide_index=True, column_config=make_column_config(warn_df))

# Compute deadlines
df["LaatsteToegestaneDag"] = (df["LeverDatum"].dt.normalize() - pd.to_timedelta(int(min_dagen_voor_lever), unit="D"))
df["VoorkeurLaatsteDag"] = (df["LeverDatum"].dt.normalize() - pd.to_timedelta(int(pref_dagen_voor_lever), unit="D"))

# Earliest due to receipt rule
def earliest_day(row):
    e = start
    if use_ontvangst and (row["_BinnenBool"] is False):
        if "AanleverDatum" in df.columns and pd.notna(row.get("AanleverDatum")):
            e = max(e, row["AanleverDatum"].normalize() + timedelta(days=int(min_dagen_ontvangst)))
    return e

df["VroegsteDag"] = df.apply(earliest_day, axis=1)

original_plan = df["PlanDatum"].dt.normalize().copy()

# ---------- Heuristic ----------
def solve_heuristic(df_in: pd.DataFrame):
    dfh = df_in.copy()
    day_m2 = {d: float(fixed_m2.get(d, 0.0)) for d in candidate_days}
    day_colors = {d: set(fixed.loc[fixed["PlanDatumN"]==d, "Kleur"].tolist()) for d in candidate_days}

    new_dates = {}
    unplanned_reason = {}

    # Sort by (deadline) then bigger M2 first (helps pack)
    dfh2 = dfh.sort_values(["LaatsteToegestaneDag", "LeverDatum", "M2"], ascending=[True, True, False])

    for idx, row in dfh2.iterrows():
        if row["Gefixeerd"] or pd.isna(row["PlanDatum"]):
            new_dates[idx] = row["PlanDatum"].normalize() if pd.notna(row["PlanDatum"]) else pd.NaT
            continue

        latest = row["LaatsteToegestaneDag"]
        earliest = row["VroegsteDag"]
        pref_latest = row["VoorkeurLaatsteDag"]

        # Candidate days for this order: within [earliest, latest]
        feas_days = [d for d in candidate_days if (d >= earliest and d <= latest)]

        if not feas_days:
            new_dates[idx] = pd.NaT
            unplanned_reason[idx] = "Niet planbaar binnen kalender/deadline"
            continue

        # Prefer days that are <= pref_latest (more buffer). Otherwise still allow up to latest.
        # Scoring: (is_after_pref, colorset_add, remaining_capacity_negative)
        best = None
        for d in feas_days:
            m2 = float(row["M2"])
            if day_m2[d] + m2 > max_m2:
                continue
            # color count check
            colors = day_colors[d]
            new_color = row["Kleur"] not in colors
            if (len(colors) + (1 if new_color else 0)) > max_kleuren:
                continue

            after_pref = 1 if d > pref_latest else 0
            # prefer not adding a new color, and prefer fuller blocks (more m2 already on that color)
            add_color = 1 if new_color else 0
            remcap = (max_m2 - (day_m2[d] + m2))  # smaller is better
            score = (after_pref, add_color, remcap)

            if best is None or score < best[0]:
                best = (score, d)

        if best is None:
            new_dates[idx] = pd.NaT
            unplanned_reason[idx] = "Capaciteit/kleur-limiet"
            continue

        d = best[1]
        new_dates[idx] = d
        day_m2[d] += float(row["M2"])
        day_colors[d].add(row["Kleur"])

    return new_dates, unplanned_reason

# ---------- CP-SAT ----------
def solve_cpsat(df_in: pd.DataFrame):
    if not ORTOOLS_OK:
        return None, {i:"OR-Tools niet beschikbaar" for i in df_in.index}

    dfi = df_in.copy()
    # Only variable orders are non-fixed and with valid dates
    var_idx = [i for i,r in dfi.iterrows() if (not r["Gefixeerd"])]
    days = candidate_days[:]  # list of pd.Timestamp
    T = len(days)

    # Map day -> t
    day_to_t = {d:i for i,d in enumerate(days)}

    # Fixed usage arrays
    fixed_m2_arr = [float(fixed_m2.get(d, 0.0)) for d in days]
    fixed_color_sets = [set(fixed.loc[fixed["PlanDatumN"]==d, "Kleur"].tolist()) for d in days]

    colors = sorted(dfi["Kleur"].dropna().astype(str).unique().tolist())
    color_to_k = {c:k for k,c in enumerate(colors)}
    K = len(colors)

    model = cp_model.CpModel()

    # Decision vars: x[i,t] assignment, plus unassigned[i]
    x = {}
    unassigned = {}
    for i in var_idx:
        unassigned[i] = model.NewBoolVar(f"unassigned_{i}")
        for t in range(T):
            x[(i,t)] = model.NewBoolVar(f"x_{i}_{t}")

        # Each order: sum x + unassigned == 1 (either placed exactly once, or unassigned)
        model.Add(sum(x[(i,t)] for t in range(T)) + unassigned[i] == 1)

    # Day capacity constraints
    for t,d in enumerate(days):
        m2_terms = []
        for i in var_idx:
            m2 = int(round(float(dfi.at[i,"M2"])))
            m2_terms.append(m2 * x[(i,t)])
        model.Add(sum(m2_terms) + int(round(fixed_m2_arr[t])) <= int(max_m2))

    # Color usage vars y[k,t]
    y = {}
    for k,c in enumerate(colors):
        for t,d in enumerate(days):
            y[(k,t)] = model.NewBoolVar(f"y_{k}_{t}")
            # If fixed already has this color, y must be 1
            if c in fixed_color_sets[t]:
                model.Add(y[(k,t)] == 1)

    # Link x -> y
    for i in var_idx:
        c = str(dfi.at[i,"Kleur"])
        k = color_to_k.get(c, None)
        if k is None:
            continue
        for t in range(T):
            model.Add(y[(k,t)] >= x[(i,t)])

    # Max colors per day
    for t in range(T):
        model.Add(sum(y[(k,t)] for k in range(K)) <= int(max_kleuren))

    # Feasibility windows (calendar + receipt + deadline)
    for i in var_idx:
        earliest = dfi.at[i,"VroegsteDag"]
        latest = dfi.at[i,"LaatsteToegestaneDag"]
        # Disallow assignment outside [earliest, latest]
        for t,d in enumerate(days):
            if d < earliest or d > latest:
                model.Add(x[(i,t)] == 0)

    # Objective: prioritize planning all orders, then minimize color-days, then preference lateness
    BIG = 1_000_000
    W_COLOR = 1000
    W_PREF = 1

    obj_terms = []
    # unassigned penalty
    obj_terms += [BIG * unassigned[i] for i in var_idx]
    # color-days penalty (encourages larger blocks / fewer colors used)
    obj_terms += [W_COLOR * y[(k,t)] for k in range(K) for t in range(T)]

    # preference penalty: if scheduled after VoorkeurLaatsteDag, add days late wrt preference
    for i in var_idx:
        pref_latest = dfi.at[i,"VoorkeurLaatsteDag"]
        for t,d in enumerate(days):
            if d > pref_latest:
                penalty = int((d - pref_latest).days)
                obj_terms.append(W_PREF * penalty * x[(i,t)])

    model.Minimize(sum(obj_terms))

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = float(time_limit)
    solver.parameters.num_search_workers = 8

    status = solver.Solve(model)

    new_dates = {}
    reasons = {}

    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        for i in var_idx:
            new_dates[i] = pd.NaT
            reasons[i] = "Geen oplossing gevonden (tijdslimiet/constraints)"
        # keep fixed as original
        for i,r in dfi.iterrows():
            if r["Gefixeerd"]:
                new_dates[i] = r["PlanDatum"].normalize() if pd.notna(r["PlanDatum"]) else pd.NaT
        return new_dates, reasons

    for i,r in dfi.iterrows():
        if r["Gefixeerd"]:
            new_dates[i] = r["PlanDatum"].normalize() if pd.notna(r["PlanDatum"]) else pd.NaT

    for i in var_idx:
        if solver.Value(unassigned[i]) == 1:
            new_dates[i] = pd.NaT
            reasons[i] = "Niet planbaar binnen constraints"
            continue
        assigned_t = None
        for t in range(T):
            if solver.Value(x[(i,t)]) == 1:
                assigned_t = t
                break
        new_dates[i] = days[assigned_t] if assigned_t is not None else pd.NaT

    return new_dates, reasons

# ---------- Run solve ----------
if optimizer.startswith("CP-SAT"):
    if not ORTOOLS_OK:
        st.error("CP-SAT gekozen, maar OR-Tools (ortools) is niet beschikbaar in deze Python omgeving. Kies 'Heuristiek (snel)' of installeer ortools.")
        st.stop()
    new_dates, reasons = solve_cpsat(df)
else:
    new_dates, reasons = solve_heuristic(df)

df["NieuwePlanDatum"] = pd.Series(new_dates)
df["OudePlanDatum"] = original_plan
df["Gewijzigd"] = (df["NieuwePlanDatum"] != df["OudePlanDatum"]) & df["NieuwePlanDatum"].notna()

df["Gefixeerd_JaNee"] = df["Gefixeerd"].map(lambda x: "Ja" if bool(x) else "Nee")

# Late list (hard deadline)
df["HardDeadlineOK"] = df["NieuwePlanDatum"].notna() & (df["NieuwePlanDatum"] <= df["LaatsteToegestaneDag"])
late = df[(df["NieuwePlanDatum"].isna()) | (~df["HardDeadlineOK"])].copy()
if len(late):
    late["Reden"] = late.index.map(lambda i: reasons.get(i, "Hard deadline overschreden"))
else:
    late["Reden"] = ""

# Herplande orders view
herpland = df[df["Gewijzigd"]].copy()
herpland["Verschil_dagen"] = (herpland["NieuwePlanDatum"] - herpland["OudePlanDatum"]).dt.days

# Summary
st.subheader("Samenvatting")
c1,c2,c3,c4 = st.columns(4)
c1.metric("Totaal orders", len(df))
c2.metric("Gefixeerd", int(df["Gefixeerd"].sum()))
c3.metric("Herpland", int(len(herpland)))
c4.metric("Te laat / niet planbaar", int(len(late)))

# Explain Binnen usage
with st.expander("Toelichting ontvangstconstraint (Binnen)"):
    st.write("**Binnen = waar**: order wordt als ontvangen beschouwd → geen aanleverdatum-buffer toegepast.")
    st.write("**Binnen = onwaar**: order nog niet ontvangen → als AanleverDatum aanwezig is, dan geldt:")
    st.write(f"- Hard: plandatum ≥ Aanleverdatum + {int(min_dagen_ontvangst)} dagen")
    st.write(f"- Soft: CP-SAT probeert plandatum ≥ Aanleverdatum + {int(pref_dagen_ontvangst)} dagen (als haalbaar)")


# ---------- Dagsamenvatting ----------
# Effective day: fixed orders keep their original PlanDatum; others use NieuwePlanDatum
df["EffectivePlanDag"] = df.apply(
    lambda r: (r["PlanDatum"].normalize() if bool(r["Gefixeerd"]) and pd.notna(r["PlanDatum"]) else r["NieuwePlanDatum"]),
    axis=1
)

day_base = df[df["EffectivePlanDag"].notna()].copy()
day_base["Dag"] = day_base["EffectivePlanDag"].dt.normalize()

# Split fixed vs newly planned (non-fixed)
fixed_part = day_base[day_base["Gefixeerd"].astype(bool)].copy()
new_part = day_base[~day_base["Gefixeerd"].astype(bool)].copy()

d_fix = (
    fixed_part.groupby("Dag")
    .agg(m2_gefixeerd=("M2","sum"), kleuren_gefixeerd=("Kleur", pd.Series.nunique), orders_gefixeerd=("OrderID","count"))
)

d_new = (
    new_part.groupby("Dag")
    .agg(m2_nieuw=("M2","sum"), kleuren_nieuw=("Kleur", pd.Series.nunique), orders_nieuw=("OrderID","count"))
)

dagsamenvatting = (
    pd.concat([d_fix, d_new], axis=1)
    .fillna(0)
    .reset_index()
)

# Totals
dagsamenvatting["Orders"] = dagsamenvatting["orders_gefixeerd"].astype(int) + dagsamenvatting["orders_nieuw"].astype(int)
dagsamenvatting["m2"] = dagsamenvatting["m2_gefixeerd"].astype(float) + dagsamenvatting["m2_nieuw"].astype(float)

# Total colors from full day_base (not sum)
d_tot_colors = day_base.groupby("Dag")["Kleur"].nunique().rename("Kleuren").reset_index()
dagsamenvatting = dagsamenvatting.merge(d_tot_colors, on="Dag", how="left")

dagsamenvatting["Over_m2"] = dagsamenvatting["m2"] > float(max_m2)
dagsamenvatting["Over_kleuren"] = dagsamenvatting["Kleuren"] > int(max_kleuren)

dagsamenvatting = dagsamenvatting.sort_values("Dag")

# Display formatting
dagsamenvatting_view = dagsamenvatting.copy()
for c in ["m2","m2_gefixeerd","m2_nieuw"]:
    dagsamenvatting_view[c] = dagsamenvatting_view[c].apply(fmt_m2)

st.subheader("Dagsamenvatting")
st.dataframe(format_dates_ddmmyyyy(dagsamenvatting_view), use_container_width=True, hide_index=True, column_config=make_column_config(dagsamenvatting_view))

# Show herplande table (planner action list)
st.subheader("Actielijst: orders met nieuwe plandatum")
if len(herpland)==0:
    st.success("Geen orders herpland.")
else:
    preferred_cols = ['Naam', 'OrderTypeAfkorting', 'OrderID', 'DatumIn', 'DeelorderID', 'VolgNr', 'PlanDatum', 'LeverDatum', 'Omschrijving', 'Kleur', 'M2', 'ReferentieKlant', 'PlanVolgordeTbvPoedercoaten', 'RALZoekcode', 'M2_Def', 'AanleverDatum', 'LeverDatumTypeAfkorting', 'Prio', 'Binnen', '_BinnenBool', 'Gefixeerd', 'LaatsteToegestaneDag', 'VoorkeurLaatsteDag', 'VroegsteDag', 'NieuwePlanDatum', 'OudePlanDatum', 'Gewijzigd', 'Gefixeerd_JaNee', 'HardDeadlineOK']
    view = herpland.copy()
    cols = [c for c in preferred_cols if c in view.columns]
    view = view[cols]  # strikt alleen deze kolommen

    # Herplande Orders: begin met vaste kolomvolgorde, vul aan met overige geselecteerde kolommen (geen dubbels)
    start_cols = ['Naam', 'OrderID', 'DeelorderID', 'VolgNr', 'Kleur', 'M2', 'Binnen', 'LeverDatum', 'LaatsteToegestaneDag', 'VroegsteDag', 'NieuwePlanDatum', 'OudePlanDatum']
    ordered = [c for c in start_cols if c in view.columns]
    for c in view.columns:
        if c not in ordered:
            ordered.append(c)
    view = view[ordered]

    view = fmt_df_m2(view, ["M2","M2_Def"])
    sort_cols = [c for c in ["NieuwePlanDatum","Kleur","OrderID"] if c in view.columns]
    if sort_cols:
        view = view.sort_values(sort_cols)
    st.dataframe(format_dates_ddmmyyyy(view), use_container_width=True, hide_index=True, column_config=make_column_config(view))

st.subheader("Detailplanning (alles)")
detail = df.copy()
detail = fmt_df_m2(detail, ["M2"])
st.dataframe(format_dates_ddmmyyyy(detail), use_container_width=True, hide_index=True, column_config=make_column_config(detail))

# Late list
st.subheader("Te laat / niet planbaar")
if len(late)==0:
    st.success("Geen te late of niet-planbare orders.")
else:
    late_view = late.copy()
    cols = ["OrderID","Kleur","M2","OudePlanDatum","NieuwePlanDatum","LeverDatum","LaatsteToegestaneDag","Gefixeerd_JaNee","Reden"]
    cols = [c for c in cols if c in late_view.columns]
    late_view = late_view[cols]
    late_view = fmt_df_m2(late_view, ["M2"])
    st.dataframe(format_dates_ddmmyyyy(late_view).sort_values(["LeverDatum"]), use_container_width=True, hide_index=True, column_config=make_column_config(late_view))

# Export

output = BytesIO()
normalize_dates(df)

with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name="Detailplanning")
    herpland_view = herpland.copy()
    # Export Herplande Orders: start cols then rest (no duplicates)
    base_cols = [c for c in ['Naam', 'OrderID', 'DeelorderID', 'VolgNr', 'Kleur', 'M2', 'Binnen', 'LeverDatum', 'LaatsteToegestaneDag', 'VroegsteDag', 'NieuwePlanDatum', 'OudePlanDatum'] if c in herpland_view.columns]
    for c in herpland_view.columns:
        if c not in base_cols:
            base_cols.append(c)
    herpland_view[base_cols].to_excel(writer, index=False, sheet_name="Herplande Orders")
    late.to_excel(writer, index=False, sheet_name="Te laat")
    if "dagsamenvatting" in globals():
        try:
            dagsamenvatting.to_excel(writer, index=False, sheet_name="Dagsamenvatting")
        except Exception:
            pass
    if "warn_df" in globals():
        try:
            warn_df.to_excel(writer, index=False, sheet_name="Fix-overload")
        except Exception:
            pass

    # Formatting: date format + column widths
    date_format = "DD-MM-YYYY"
    caps = {"ReferentieKlant": 35, "Omschrijving": 35}
    for wsname in writer.book.sheetnames:
        ws = writer.book[wsname]
        header_row = [cell.value for cell in ws[1]]
        for j, header in enumerate(header_row, start=1):
            if header is None:
                continue
            h = str(header)
            if ("Datum" in h) or h.endswith("Dag"):
                for row in ws.iter_rows(min_row=2, min_col=j, max_col=j):
                    cell = row[0]
                    if cell.value is None:
                        continue
                    if hasattr(cell.value, "year"):
                        cell.number_format = date_format
        # autosize columns
        for col_cells in ws.columns:
            col_letter = col_cells[0].column_letter
            head = str(col_cells[0].value) if col_cells and col_cells[0].value is not None else ""
            max_len = len(head)
            for cell in col_cells[1:]:
                if cell.value is None:
                    continue
                s = str(cell.value)
                if len(s) > max_len:
                    max_len = len(s)
            width = max(10, min(60, max_len + 2))
            if head in caps:
                width = min(width, caps[head])
            ws.column_dimensions[col_letter].width = width

output.seek(0)

st.download_button("Download resultaat (Excel)", data=output.getvalue(), file_name="planning_resultaat.xlsx")
