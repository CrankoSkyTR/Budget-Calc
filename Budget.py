import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
from io import BytesIO

st.set_page_config(page_title="Ops Budget", layout="wide")

ACCOUNTS = [
    "Sky X Chat",
    "SALT IB",
    "Inditex",
    "BackMarket",
    "EMMA",
    "Chrono24",
    "Adidas",
    "TP Vision",
]
DEFAULT_LANGUAGES = ["DE", "FR", "IT", "TR", "EN",
                     "ES", "NL", "AR", "CH/DE", "DE/EN", "TR/EN"]
ROLES = ["Operation Manager", "Teamleader", "Trainer & Quality",
         "RTA", "Planner", "WFM", "Operation Support"]


# -----------------------------
# Helpers
# -----------------------------
def month_range(start_ym: str, months: int) -> list:
    start = pd.to_datetime(start_ym + "-01")
    return pd.date_range(start=start, periods=months, freq="MS").strftime("%Y-%m").tolist()


def df_to_excel_bytes(df_dict: dict) -> bytes:
    output = BytesIO()
    try:
        engine = "xlsxwriter"
        __import__("xlsxwriter")
    except Exception:
        engine = "openpyxl"

    with pd.ExcelWriter(output, engine=engine) as writer:
        for sheet, df in df_dict.items():
            df.to_excel(writer, sheet_name=sheet[:31], index=False)
    return output.getvalue()


def money_fmt(s: pd.Series, decimals: int = 2) -> pd.Series:
    return pd.to_numeric(s, errors="coerce").fillna(0.0).map(lambda x: f"{x:,.{decimals}f}")


def gm_pct_from_series(gm: pd.Series, revenue: pd.Series) -> pd.Series:
    gm = pd.to_numeric(gm, errors="coerce").fillna(0.0)
    revenue = pd.to_numeric(revenue, errors="coerce").fillna(0.0)
    out = np.where(revenue != 0, gm / revenue, np.where(gm < 0, -1.0, 0.0))
    return pd.Series(out, index=gm.index)


def read_optional_sheet(xls: pd.ExcelFile, sheet_name: str, columns: list) -> pd.DataFrame:
    if sheet_name not in xls.sheet_names:
        return pd.DataFrame(columns=columns)
    df = pd.read_excel(xls, sheet_name=sheet_name)
    for c in columns:
        if c not in df.columns:
            df[c] = np.nan
    return df[columns].copy()


def normalize_language_key(v) -> str:
    s = str(v).strip()
    if " - " in s:
        s = s.split(" - ", 1)[0].strip()
    return s


def normalize_plan(plan_df: pd.DataFrame, months: list, combos_df: pd.DataFrame) -> pd.DataFrame:
    base = combos_df[["Account", "Language"]].drop_duplicates().copy()
    if base.empty:
        return pd.DataFrame(
            columns=["Month", "Account", "Language",
                     "Production_Hours", "FTE", "Notes"]
        )

    target = []
    for m in months:
        tmp = base.copy()
        tmp["Month"] = m
        target.append(tmp)
    target_df = pd.concat(target, ignore_index=True)

    keep_cols = ["Month", "Account", "Language",
                 "Production_Hours", "FTE", "Notes"]
    if plan_df.empty:
        out = target_df.copy()
        out["Production_Hours"] = 0.0
        out["FTE"] = 0.0
        out["Notes"] = ""
        return out[keep_cols]

    plan_df = plan_df.copy()
    for c in keep_cols:
        if c not in plan_df.columns:
            plan_df[c] = 0.0 if c in ["Production_Hours", "FTE"] else ""

    merged = target_df.merge(
        plan_df[keep_cols],
        on=["Month", "Account", "Language"],
        how="left",
    )
    merged["Production_Hours"] = pd.to_numeric(
        merged["Production_Hours"], errors="coerce").fillna(0.0)
    merged["FTE"] = pd.to_numeric(merged["FTE"], errors="coerce").fillna(0.0)
    merged["Notes"] = merged["Notes"].fillna("")
    return merged[keep_cols]


def workdays_in_month(month_ym: str) -> int:
    start = pd.to_datetime(month_ym + "-01", errors="coerce")
    if pd.isna(start):
        return 0
    end = start + pd.offsets.MonthEnd(0)
    return len(pd.date_range(start=start, end=end, freq="B"))


def aggregate_overhead_monthly(overhead_df: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    if overhead_df is None or overhead_df.empty:
        return pd.DataFrame(columns=["Account", "Monthly_Overhead_TRY"])

    oh = overhead_df.copy()
    # Detailed mode: Account + Role + FTE + salary components
    for col, default in [
        ("FTE", 0.0),
        ("Salary", 0.0),
        ("OSS", cfg["default_oss"]),
        ("Food", cfg["default_food"]),
        ("Goalpex_%", cfg["default_goalpex"]),
        ("Additional_Cost", 0.0),
        ("Transport_Allowance_Net", 0.0),
        ("Sales_Bonus_Net", 0.0),
    ]:
        if col not in oh.columns:
            oh[col] = default
        oh[col] = pd.to_numeric(oh[col], errors="coerce").fillna(default)

    if "Monthly_Overhead_TRY" in oh.columns:
        manual = pd.to_numeric(oh["Monthly_Overhead_TRY"], errors="coerce")
    else:
        manual = pd.Series(np.nan, index=oh.index, dtype="float64")
    oh["Goalpex"] = oh["Salary"] * oh["Goalpex_%"]
    oh["Net_Per_FTE"] = oh["Salary"] + oh["Goalpex"] + \
        oh["OSS"] + oh["Food"] + oh["Additional_Cost"] + \
        oh["Transport_Allowance_Net"] + oh["Sales_Bonus_Net"]
    computed = oh["Net_Per_FTE"] * cfg["brut_multiplier"] * oh["FTE"]
    oh["Monthly_Overhead_TRY"] = manual.combine_first(computed).fillna(0.0)
    return oh.groupby("Account", as_index=False)["Monthly_Overhead_TRY"].sum()


def clean_overhead_rows(df: pd.DataFrame, default_account: str, cfg: dict) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=["Account", "Role", "FTE", "Salary", "OSS", "Food", "Goalpex_%", "Additional_Cost", "Transport_Allowance_Net", "Sales_Bonus_Net"])
    out = df.copy()
    for c, default in [
        ("Account", default_account),
        ("Role", ROLES[0]),
        ("FTE", 1.0),
        ("Salary", 0.0),
        ("OSS", cfg["default_oss"]),
        ("Food", cfg["default_food"]),
        ("Goalpex_%", cfg["default_goalpex"]),
        ("Additional_Cost", 0.0),
        ("Transport_Allowance_Net", 0.0),
        ("Sales_Bonus_Net", 0.0),
    ]:
        if c not in out.columns:
            out[c] = default
    out["Account"] = out["Account"].fillna(default_account).astype(str)
    out["Role"] = out["Role"].fillna(ROLES[0]).astype(str)
    for c, default in [("FTE", 1.0), ("Salary", 0.0), ("OSS", cfg["default_oss"]), ("Food", cfg["default_food"]), ("Goalpex_%", cfg["default_goalpex"]), ("Additional_Cost", 0.0), ("Transport_Allowance_Net", 0.0), ("Sales_Bonus_Net", 0.0)]:
        out[c] = pd.to_numeric(out[c], errors="coerce").fillna(default)
    out = out[out["Account"].isin(ACCOUNTS)]
    out = out[out["Role"].isin(ROLES)]
    return out[["Account", "Role", "FTE", "Salary", "OSS", "Food", "Goalpex_%", "Additional_Cost", "Transport_Allowance_Net", "Sales_Bonus_Net"]].copy()


def compute_budget(plan_df, prices_df, cost_df, fx_df, overhead_df, cfg):
    if plan_df.empty:
        return pd.DataFrame()

    # Keep plan input clean to avoid merge suffix collisions from stale columns
    plan_cols = ["Month", "Account", "Language", "Production_Hours", "FTE", "Notes"]
    df = plan_df.copy()
    for c in plan_cols:
        if c not in df.columns:
            df[c] = "" if c in ["Month", "Account", "Language", "Notes"] else 0.0
    df = df[plan_cols].copy()

    # Unit price + billing mode + currency by operation/language
    p = prices_df.copy()
    for c, default in [("Account", ""), ("Language", ""), ("UnitPrice", 0.0), ("Currency", "TRY"), ("Billing_Mode", "Unit Price × Production Hours"), ("Note", "")]:
        if c not in p.columns:
            p[c] = default
    p = p[["Account", "Language", "UnitPrice", "Currency", "Billing_Mode", "Note"]].copy()
    p["Account"] = p["Account"].astype(str)
    p["Language"] = p["Language"].astype(str)
    p = p.drop_duplicates(subset=["Account", "Language"], keep="last")
    p["UnitPrice"] = pd.to_numeric(p["UnitPrice"], errors="coerce").fillna(0.0)
    df = df.merge(p, on=["Account", "Language"], how="left")
    df["UnitPrice"] = pd.to_numeric(
        df["UnitPrice"], errors="coerce").fillna(0.0)
    df["Billing_Mode"] = df["Billing_Mode"].fillna(
        "Unit Price × Production Hours")
    df["Currency"] = df["Currency"].fillna("TRY")

    # FX by Month + Currency (TRY defaults to 1)
    fx = fx_df.copy()
    for c, default in [("Month", ""), ("Currency", "TRY"), ("FX_Rate", 1.0)]:
        if c not in fx.columns:
            fx[c] = default
    fx = fx[["Month", "Currency", "FX_Rate"]].copy()
    fx = fx.drop_duplicates(subset=["Month", "Currency"], keep="last")
    if not fx.empty:
        fx["FX_Rate"] = pd.to_numeric(
            fx["FX_Rate"], errors="coerce").fillna(1.0)
        df = df.merge(fx, on=["Month", "Currency"], how="left")
    else:
        df["FX_Rate"] = np.nan
    df["FX_Rate"] = np.where(df["Currency"] == "TRY",
                             1.0, pd.to_numeric(df["FX_Rate"], errors="coerce"))
    df["FX_Rate"] = pd.to_numeric(df["FX_Rate"], errors="coerce").fillna(1.0)

    # COLA application by month
    mdt = pd.to_datetime(df["Month"] + "-01", errors="coerce")
    cola_start = pd.to_datetime(
        cfg["cola_start_month"] + "-01", errors="coerce")
    df["COLA_%"] = np.where(mdt >= cola_start, cfg["cola_pct"], 0.0)

    # Cost defaults by (Account, Language), fallback to (All, Language)
    cost_cols = [
        "Account",
        "Language",
        "Salary",
        "OSS",
        "Food",
        "Goalpex_%",
        "Additional_Cost",
        "Transport_Allowance_Net",
        "Sales_Bonus_Net",
    ]
    c = cost_df.copy()
    for cc in cost_cols:
        if cc not in c.columns:
            c[cc] = np.nan
    c = c[cost_cols].copy()
    df["Language_Key"] = df["Language"].apply(normalize_language_key)
    c["Language_Key"] = c["Language"].apply(normalize_language_key)
    for col in ["Salary", "OSS", "Food", "Goalpex_%", "Additional_Cost", "Transport_Allowance_Net", "Sales_Bonus_Net"]:
        c[col] = pd.to_numeric(c[col], errors="coerce")

    val_cols = ["Salary", "OSS", "Food", "Goalpex_%", "Additional_Cost", "Transport_Allowance_Net", "Sales_Bonus_Net"]

    spec = c[c["Account"] != "All"][["Account", "Language"] + val_cols].drop_duplicates(
        subset=["Account", "Language"], keep="last"
    ).copy()
    all_lang = c[c["Account"] == "All"][["Language"] + val_cols].drop_duplicates(
        subset=["Language"], keep="last"
    ).copy()

    df = df.merge(
        spec,
        on=["Account", "Language"],
        how="left",
        suffixes=("", "_spec"),
    )
    df = df.merge(
        all_lang,
        on=["Language"],
        how="left",
        suffixes=("_spec", "_all"),
    )

    spec_key = c[c["Account"] != "All"][["Account", "Language_Key"] + val_cols].drop_duplicates(
        subset=["Account", "Language_Key"], keep="last"
    ).copy()
    all_key = c[c["Account"] == "All"][["Language_Key"] + val_cols].drop_duplicates(
        subset=["Language_Key"], keep="last"
    ).copy()
    df = df.merge(
        spec_key,
        on=["Account", "Language_Key"],
        how="left",
        suffixes=("", "_spec_key"),
    )
    df = df.merge(
        all_key,
        on=["Language_Key"],
        how="left",
        suffixes=("_spec_key", "_all_key"),
    )

    def pick(col, default):
        s = pd.to_numeric(df.get(f"{col}_spec"), errors="coerce")
        a = pd.to_numeric(df.get(f"{col}_all"), errors="coerce")
        sk = pd.to_numeric(df.get(f"{col}_spec_key"), errors="coerce")
        ak = pd.to_numeric(df.get(f"{col}_all_key"), errors="coerce")
        return s.combine_first(a).combine_first(sk).combine_first(ak).fillna(default)

    df["Salary"] = pick("Salary", 0.0)
    df["OSS"] = pick("OSS", cfg["default_oss"])
    df["Food"] = pick("Food", cfg["default_food"])
    df["Goalpex_%"] = pick("Goalpex_%", cfg["default_goalpex"])
    df["Additional_Cost"] = pick("Additional_Cost", 0.0)
    df["Transport_Allowance_Net"] = pick("Transport_Allowance_Net", 0.0)
    df["Sales_Bonus_Net"] = pick("Sales_Bonus_Net", 0.0)

    df["FTE"] = pd.to_numeric(df["FTE"], errors="coerce").fillna(0.0)
    df["Production_Hours"] = pd.to_numeric(
        df["Production_Hours"], errors="coerce").fillna(0.0)

    # Agent cost (TRY)
    df["Goalpex"] = df["Salary"] * df["Goalpex_%"]
    df["Net_Per_FTE"] = df["Salary"] + df["Goalpex"] + \
        df["OSS"] + df["Food"] + df["Additional_Cost"] + \
        df["Transport_Allowance_Net"] + df["Sales_Bonus_Net"]
    df["Agent_Cost"] = df["Net_Per_FTE"] * cfg["brut_multiplier"] * df["FTE"]

    # Overhead cost by account
    oh = aggregate_overhead_monthly(overhead_df, cfg)
    if not oh.empty:
        df = df.merge(oh[["Account", "Monthly_Overhead_TRY"]],
                      on="Account", how="left")
    else:
        df["Monthly_Overhead_TRY"] = 0.0
    df["Monthly_Overhead_TRY"] = pd.to_numeric(
        df["Monthly_Overhead_TRY"], errors="coerce").fillna(0.0)

    rows_per_month_acc = df.groupby(["Month", "Account"])[
        "Account"].transform("count").replace(0, 1)
    df["Overhead_Cost"] = df["Monthly_Overhead_TRY"] / rows_per_month_acc

    df["Total Cost"] = df["Agent_Cost"] + df["Overhead_Cost"]

    # Revenue
    shrink = cfg["shrinkage_pct"]
    df["Eff_Production_Hours"] = df["Production_Hours"] * (1 - shrink)
    df["Eff_FTE"] = df["FTE"] * (1 - shrink)
    df["Eff_FTE_Hours"] = df["Eff_FTE"] * 180.0

    df["Billable_Hours"] = np.where(
        df["Billing_Mode"] == "Unit Price × FTE",
        df["Eff_FTE_Hours"],
        df["Eff_Production_Hours"],
    )

    df["Adj. Unit Price"] = df["UnitPrice"] * (1 + df["COLA_%"])
    df["Revenue"] = df["Billable_Hours"] * \
        df["Adj. Unit Price"] * df["FX_Rate"]
    df["GM"] = df["Revenue"] - df["Total Cost"]
    df["GM_%"] = np.where(df["Revenue"] == 0, 0.0, df["GM"] / df["Revenue"])

    # Keep clean columns
    show_cols = [
        "Month", "Account", "Language", "Production_Hours", "FTE",
        "Agent_Cost", "Overhead_Cost", "Total Cost",
        "UnitPrice", "COLA_%", "Adj. Unit Price", "Billing_Mode", "Currency", "FX_Rate",
        "Eff_Production_Hours", "Eff_FTE", "Eff_FTE_Hours", "Billable_Hours",
        "Revenue", "GM", "GM_%", "Notes"
    ]
    show_cols = [c for c in show_cols if c in df.columns]
    return df[show_cols].copy()


# -----------------------------
# Session defaults
# -----------------------------
if "unit_prices_df" not in st.session_state:
    st.session_state.unit_prices_df = pd.DataFrame(
        [{"Account": a, "Language": l, "UnitPrice": 0.0, "Currency": "TRY", "Billing_Mode": "Unit Price × Production Hours", "Note": ""}
         for a in ACCOUNTS for l in ["DE", "FR", "EN"]]
    )
if "Note" not in st.session_state.unit_prices_df.columns:
    st.session_state.unit_prices_df["Note"] = ""
st.session_state.unit_prices_df["Note"] = st.session_state.unit_prices_df["Note"].fillna("").astype(str)

if "cost_defaults_df" not in st.session_state:
    st.session_state.cost_defaults_df = pd.DataFrame(
        [{"Account": "All", "Language": l, "Salary": 0.0, "OSS": 2083.0, "Food": 5850.0, "Goalpex_%": 0.10, "Additional_Cost": 0.0, "Transport_Allowance_Net": 0.0, "Sales_Bonus_Net": 0.0}
         for l in ["DE", "FR", "EN"]]
    )

if "overhead_df" not in st.session_state:
    st.session_state.overhead_df = pd.DataFrame(
        [
            {
                "Account": ACCOUNTS[0],
                "Role": "Operation Manager",
                "FTE": 1.0,
                "Salary": 0.0,
                "OSS": 2083.0,
                "Food": 5850.0,
                "Goalpex_%": 0.10,
                "Additional_Cost": 0.0,
                "Transport_Allowance_Net": 0.0,
                "Sales_Bonus_Net": 0.0,
            }
        ]
    )

if "plan_df" not in st.session_state:
    st.session_state.plan_df = pd.DataFrame(
        columns=["Month", "Account", "Language", "Production_Hours", "FTE", "Notes"])


# -----------------------------
# Sidebar controls
# -----------------------------
st.title("Operations Budget Calculator")

with st.sidebar:
    st.header("Setup")
    selected_accounts = st.multiselect(
        "Operations", ACCOUNTS, default=["Sky X Chat"])
    if not selected_accounts:
        selected_accounts = ["Sky X Chat"]

    start_month = st.text_input("Start Month (YYYY-MM)", "2026-02")
    n_months = st.number_input(
        "Number of months", min_value=1, max_value=60, value=12)
    months = month_range(start_month, int(n_months))

    st.divider()
    st.subheader("Assumptions")
    cola_pct = st.number_input(
        "COLA % (decimal)", value=0.0, step=0.01, format="%.4f")
    cola_start_month = st.selectbox(
        "Apply COLA starting month", options=months, index=0)
    brut_multiplier = st.slider(
        "Brut Multiplier", min_value=1.0, max_value=3.0, value=1.58, step=0.01)
    shrinkage_pct = st.slider(
        "Shrinkage %", min_value=0.0, max_value=100.0, value=0.0, step=0.5) / 100.0
    hours_per_workday = st.number_input(
        "Hours per Workday (Capacity)", min_value=0.0, max_value=24.0, value=7.75, step=0.25, format="%.2f")


cfg = {
    "cola_pct": cola_pct,
    "cola_start_month": cola_start_month,
    "brut_multiplier": brut_multiplier,
    "shrinkage_pct": shrinkage_pct,
    "default_oss": 2083.0,
    "default_food": 5850.0,
    "default_goalpex": 0.10,
    "hours_per_workday": hours_per_workday,
}


# -----------------------------
# Tabs
# -----------------------------
tab_setup, tab_plan, tab_results, tab_capacity = st.tabs(
    ["1) Setup", "2) Monthly Plan", "3) Results", "4) Capacity"])

with tab_setup:
    st.subheader("Required Inputs")
    st.caption(
        "You can import a previously downloaded project file to reuse all inputs.")

    st.markdown("**Project File (load previous inputs)**")
    project_file = st.file_uploader("Load project XLSX", type=[
                                    "xlsx"], key="project_loader")
    if project_file is not None:
        try:
            xls = pd.ExcelFile(project_file)
            up_cols = ["Account", "Language",
                       "UnitPrice", "Currency", "Billing_Mode", "Note"]
            cd_cols = ["Account", "Language", "Salary",
                       "OSS", "Food", "Goalpex_%", "Additional_Cost", "Transport_Allowance_Net", "Sales_Bonus_Net"]
            oh_cols = ["Account", "Role", "FTE", "Salary",
                       "OSS", "Food", "Goalpex_%", "Additional_Cost", "Transport_Allowance_Net", "Sales_Bonus_Net"]
            plan_cols = ["Month", "Account", "Language",
                         "Production_Hours", "FTE", "Notes"]
            fx_cols = ["Month", "Currency", "FX_Rate"]

            loaded_up = read_optional_sheet(xls, "Input_UnitPrices", up_cols)
            loaded_cd = read_optional_sheet(xls, "Input_CostDefaults", cd_cols)
            loaded_oh = read_optional_sheet(xls, "Input_Overhead", oh_cols)
            loaded_plan = read_optional_sheet(xls, "Input_Plan", plan_cols)
            loaded_fx = read_optional_sheet(xls, "Input_FX", fx_cols)

            if not loaded_up.empty:
                st.session_state.unit_prices_df = loaded_up
                if "Note" not in st.session_state.unit_prices_df.columns:
                    st.session_state.unit_prices_df["Note"] = ""
                st.session_state.unit_prices_df["Note"] = st.session_state.unit_prices_df["Note"].fillna("").astype(str)
            if not loaded_cd.empty:
                st.session_state.cost_defaults_df = loaded_cd
            if not loaded_oh.empty:
                st.session_state.overhead_df = loaded_oh
            if not loaded_plan.empty:
                st.session_state.plan_df = loaded_plan
            if not loaded_fx.empty:
                st.session_state.fx_rates_df = loaded_fx
            st.success("Project inputs loaded.")
        except Exception as e:
            st.error(f"Could not load project file: {e}")

    st.markdown("**Unit Prices (per Operation + Language)**")
    up_view = st.session_state.unit_prices_df.copy()
    if "Note" not in up_view.columns:
        up_view["Note"] = ""
    up_view["Note"] = up_view["Note"].fillna("").astype(str)
    up_view = up_view[up_view["Account"].isin(selected_accounts)]
    with st.form("unit_prices_form"):
        up_edit = st.data_editor(
            up_view,
            num_rows="dynamic",
            use_container_width=True,
            key="unit_prices_editor",
            column_config={
                "Account": st.column_config.SelectboxColumn("Account", options=ACCOUNTS),
                "Language": st.column_config.TextColumn("Language"),
                "UnitPrice": st.column_config.NumberColumn("Unit Price", min_value=0.0, step=0.01, format="%.2f"),
                "Currency": st.column_config.SelectboxColumn("Currency", options=["TRY", "EUR", "USD"]),
                "Billing_Mode": st.column_config.SelectboxColumn("Billing Mode", options=["Unit Price × Production Hours", "Unit Price × FTE"]),
                "Note": st.column_config.TextColumn("Note"),
            },
        )
        save_up = st.form_submit_button("Save Unit Prices")
    if save_up:
        if "Note" not in up_edit.columns:
            up_edit["Note"] = ""
        full_up = st.session_state.unit_prices_df.copy()
        full_up = full_up[~full_up["Account"].isin(selected_accounts)]
        st.session_state.unit_prices_df = pd.concat(
            [full_up, up_edit], ignore_index=True)
        st.session_state.unit_prices_df["Language"] = st.session_state.unit_prices_df["Language"].astype(
            str).str.strip()
        st.session_state.unit_prices_df["Note"] = st.session_state.unit_prices_df["Note"].fillna("").astype(str)

    st.markdown(
        "**Operation Cost Defaults (per Operation + Language, fallback via Account='All')**")
    cd_view = st.session_state.cost_defaults_df.copy()
    cd_view = cd_view[cd_view["Account"].isin(["All"] + selected_accounts)]
    with st.form("cost_defaults_form"):
        cd_edit = st.data_editor(
            cd_view,
            num_rows="dynamic",
            use_container_width=True,
            key="cost_defaults_editor",
            column_config={
                "Account": st.column_config.SelectboxColumn("Account", options=["All"] + ACCOUNTS),
                "Language": st.column_config.SelectboxColumn("Language", options=DEFAULT_LANGUAGES),
                "Salary": st.column_config.NumberColumn("Salary (TRY)", min_value=0.0, step=100.0, format="%.2f"),
                "OSS": st.column_config.NumberColumn("OSS (TRY)", min_value=0.0, step=10.0, format="%.2f"),
                "Food": st.column_config.NumberColumn("Food (TRY)", min_value=0.0, step=10.0, format="%.2f"),
                "Goalpex_%": st.column_config.NumberColumn("Goalpex %", min_value=0.0, step=0.01, format="%.2f"),
                "Additional_Cost": st.column_config.NumberColumn("Additional (TRY)", min_value=0.0, step=10.0, format="%.2f"),
                "Transport_Allowance_Net": st.column_config.NumberColumn("Transportation Allowance (Net)", min_value=0.0, step=10.0, format="%.2f"),
                "Sales_Bonus_Net": st.column_config.NumberColumn("Sales Bonus (Net)", min_value=0.0, step=10.0, format="%.2f"),
            },
        )
        save_cd = st.form_submit_button("Save Operation Costs")
    if save_cd:
        full_cd = st.session_state.cost_defaults_df.copy()
        full_cd = full_cd[~full_cd["Account"].isin(
            ["All"] + selected_accounts)]
        st.session_state.cost_defaults_df = pd.concat(
            [full_cd, cd_edit], ignore_index=True)
        st.session_state.cost_defaults_df["Language"] = st.session_state.cost_defaults_df["Language"].astype(
            str).str.strip()

    st.markdown("**Monthly Overhead (optional, by Operation + Role)**")
    oh_view = clean_overhead_rows(
        st.session_state.overhead_df.copy(),
        default_account=selected_accounts[0],
        cfg=cfg,
    )
    oh_view = oh_view[oh_view["Account"].isin(selected_accounts)]
    with st.form("overhead_form"):
        oh_edit = st.data_editor(
            oh_view,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            key="overhead_editor",
            column_config={
                "Account": st.column_config.SelectboxColumn("Account", options=ACCOUNTS),
                "Role": st.column_config.SelectboxColumn("Role", options=ROLES),
                "FTE": st.column_config.NumberColumn("FTE", min_value=0.0, step=0.1, format="%.2f"),
                "Salary": st.column_config.NumberColumn("Salary (TRY)", min_value=0.0, step=100.0, format="%.2f"),
                "OSS": st.column_config.NumberColumn("OSS (TRY)", min_value=0.0, step=10.0, format="%.2f"),
                "Food": st.column_config.NumberColumn("Food (TRY)", min_value=0.0, step=10.0, format="%.2f"),
                "Goalpex_%": st.column_config.NumberColumn("Goalpex %", min_value=0.0, step=0.01, format="%.2f"),
                "Additional_Cost": st.column_config.NumberColumn("Additional (TRY)", min_value=0.0, step=10.0, format="%.2f"),
                "Transport_Allowance_Net": st.column_config.NumberColumn("Transportation Allowance (Net)", min_value=0.0, step=10.0, format="%.2f"),
                "Sales_Bonus_Net": st.column_config.NumberColumn("Sales Bonus (Net)", min_value=0.0, step=10.0, format="%.2f"),
            },
        )
        save_oh = st.form_submit_button("Save Overhead")
    if save_oh:
        oh_edit = clean_overhead_rows(
            oh_edit,
            default_account=selected_accounts[0],
            cfg=cfg,
        )
        full_oh = st.session_state.overhead_df.copy()
        full_oh = full_oh[~full_oh["Account"].isin(selected_accounts)]
        st.session_state.overhead_df = pd.concat(
            [full_oh, oh_edit], ignore_index=True)

    # Show monthly overhead summary from detailed rows
    oh_sum = aggregate_overhead_monthly(
        st.session_state.overhead_df[st.session_state.overhead_df["Account"].isin(
            selected_accounts)],
        cfg,
    )
    if not oh_sum.empty:
        oh_sum_show = oh_sum.copy()
        oh_sum_show["Monthly_Overhead_TRY"] = money_fmt(
            oh_sum_show["Monthly_Overhead_TRY"])
        st.caption("Calculated Monthly Overhead by Operation")
        st.dataframe(oh_sum_show, use_container_width=True)

    st.markdown("**FX Rates (for EUR/USD rows in Unit Price)**")
    used_ccy = st.session_state.unit_prices_df[
        st.session_state.unit_prices_df["Account"].isin(selected_accounts)
    ]["Currency"].dropna().unique().tolist()
    used_ccy = [c for c in used_ccy if c in ["EUR", "USD"]]

    fx_seed = pd.DataFrame(
        [{"Month": m, "Currency": c, "FX_Rate": 1.0}
            for m in months for c in used_ccy]
    )
    if fx_seed.empty:
        fx_seed = pd.DataFrame(columns=["Month", "Currency", "FX_Rate"])
    if "fx_rates_df" in st.session_state and not st.session_state.fx_rates_df.empty:
        current_fx = st.session_state.fx_rates_df.copy()
        if {"Month", "Currency", "FX_Rate"}.issubset(set(current_fx.columns)) and not fx_seed.empty:
            fx_merged = fx_seed.merge(
                current_fx, on=["Month", "Currency"], how="left", suffixes=("", "_old"))
            fx_merged["FX_Rate"] = pd.to_numeric(
                fx_merged["FX_Rate_old"], errors="coerce").fillna(fx_merged["FX_Rate"])
            fx_seed = fx_merged[["Month", "Currency", "FX_Rate"]]

    if not fx_seed.empty:
        with st.form("fx_form"):
            fx_edit = st.data_editor(
                fx_seed.sort_values(["Month", "Currency"]),
                num_rows="dynamic",
                use_container_width=True,
                key="fx_editor",
                column_config={
                    "Month": st.column_config.TextColumn("Month"),
                    "Currency": st.column_config.SelectboxColumn("Currency", options=["EUR", "USD"]),
                    "FX_Rate": st.column_config.NumberColumn("TRY per Currency", min_value=0.0, step=0.01, format="%.4f"),
                },
            )
            save_fx = st.form_submit_button("Save FX")
        if save_fx:
            st.session_state.fx_rates_df = fx_edit
    else:
        st.info("No EUR/USD in Unit Prices. FX table is not required.")
        st.session_state.fx_rates_df = pd.DataFrame(
            columns=["Month", "Currency", "FX_Rate"])

with tab_plan:
    st.subheader("Monthly Plan")
    st.caption(
        "Just enter FTE and/or Production Hours. Rows are auto-generated from Setup (Operation + Language + Months).")

    combos = st.session_state.unit_prices_df.copy()
    combos = combos[combos["Account"].isin(selected_accounts)]
    combos = combos[["Account", "Language"]].drop_duplicates()
    combos["Language"] = combos["Language"].astype(str).str.strip()
    combos = combos[combos["Language"] != ""]

    st.session_state.plan_df = normalize_plan(
        st.session_state.plan_df, months, combos)

    template = st.session_state.plan_df.copy()
    xlsx_template = df_to_excel_bytes({"Plan_Template": template})
    st.download_button(
        "⬇️ Download Plan Template (xlsx)",
        data=xlsx_template,
        file_name="plan_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    up = st.file_uploader("Upload filled Plan Template (xlsx)", type=["xlsx"])
    if up is not None:
        try:
            up_df = pd.read_excel(up)
            req = {"Month", "Account", "Language"}
            if not req.issubset(set(up_df.columns)):
                st.error("Template must include Month, Account, Language.")
            else:
                m = st.session_state.plan_df.merge(
                    up_df,
                    on=["Month", "Account", "Language"],
                    how="left",
                    suffixes=("", "_new"),
                )
                for col in ["FTE", "Production_Hours", "Notes"]:
                    if f"{col}_new" in m.columns:
                        m[col] = m[f"{col}_new"].combine_first(m[col])
                        m = m.drop(columns=[f"{col}_new"])
                st.session_state.plan_df = m
                st.success("Plan imported.")
        except Exception as e:
            st.error(f"Could not read file: {e}")

    with st.form("plan_form"):
        plan_edit = st.data_editor(
            st.session_state.plan_df,
            num_rows="dynamic",
            use_container_width=True,
            key="plan_editor",
            column_config={
                "Month": st.column_config.TextColumn("Month"),
                "Account": st.column_config.SelectboxColumn("Account", options=selected_accounts),
                "Language": st.column_config.TextColumn("Language"),
                "Production_Hours": st.column_config.NumberColumn("Production Hours", min_value=0.0, step=1.0, format="%.2f"),
                "FTE": st.column_config.NumberColumn("FTE", min_value=0.0, step=0.1, format="%.2f"),
                "Notes": st.column_config.TextColumn("Notes"),
            },
        )
        save_plan = st.form_submit_button("Save Monthly Plan")
    if save_plan:
        st.session_state.plan_df = plan_edit

with tab_results:
    st.subheader("Results")

    base_result = compute_budget(
        st.session_state.plan_df,
        st.session_state.unit_prices_df,
        st.session_state.cost_defaults_df,
        st.session_state.fx_rates_df,
        st.session_state.overhead_df,
        cfg,
    )

    if base_result.empty:
        st.info("No calculable rows yet. Please fill Setup + Monthly Plan.")
    else:
        # Scenario Compare (Base vs What-if)
        with st.expander("Scenario Compare (Base vs What-if)", expanded=False):
            enable_whatif = st.checkbox("Enable What-if Scenario", value=False)
            wc1, wc2, wc3 = st.columns(3)
            with wc1:
                wf_cola = st.number_input(
                    "What-if COLA % (decimal)", value=float(cfg["cola_pct"]), step=0.01, format="%.4f")
                wf_brut = st.number_input(
                    "What-if Brut Multiplier", value=float(cfg["brut_multiplier"]), step=0.01, format="%.2f")
            with wc2:
                wf_shrink = st.number_input("What-if Shrinkage %", value=float(
                    cfg["shrinkage_pct"] * 100), step=0.5, format="%.2f") / 100.0
                wf_price_mult = st.number_input(
                    "What-if Unit Price Multiplier", value=1.0, step=0.01, format="%.3f")
            with wc3:
                wf_salary_mult = st.number_input(
                    "What-if Salary Multiplier", value=1.0, step=0.01, format="%.3f")
                wf_overhead_mult = st.number_input(
                    "What-if Overhead Multiplier", value=1.0, step=0.01, format="%.3f")

        result = base_result.copy()
        whatif_result = None
        if enable_whatif:
            wf_cfg = cfg.copy()
            wf_cfg["cola_pct"] = wf_cola
            wf_cfg["brut_multiplier"] = wf_brut
            wf_cfg["shrinkage_pct"] = wf_shrink

            wf_prices = st.session_state.unit_prices_df.copy()
            wf_prices["UnitPrice"] = pd.to_numeric(
                wf_prices["UnitPrice"], errors="coerce").fillna(0.0) * wf_price_mult

            wf_cost = st.session_state.cost_defaults_df.copy()
            wf_cost["Salary"] = pd.to_numeric(
                wf_cost["Salary"], errors="coerce").fillna(0.0) * wf_salary_mult

            wf_overhead = st.session_state.overhead_df.copy()
            if "Monthly_Overhead_TRY" in wf_overhead.columns:
                wf_overhead["Monthly_Overhead_TRY"] = pd.to_numeric(
                    wf_overhead["Monthly_Overhead_TRY"], errors="coerce"
                ).fillna(0.0) * wf_overhead_mult
            for c in ["Salary", "OSS", "Food", "Additional_Cost", "Transport_Allowance_Net", "Sales_Bonus_Net"]:
                if c in wf_overhead.columns:
                    wf_overhead[c] = pd.to_numeric(
                        wf_overhead[c], errors="coerce").fillna(0.0) * wf_overhead_mult

            whatif_result = compute_budget(
                st.session_state.plan_df,
                wf_prices,
                wf_cost,
                st.session_state.fx_rates_df,
                wf_overhead,
                wf_cfg,
            )

        total_revenue = result["Revenue"].sum()
        total_agent = result["Agent_Cost"].sum()
        total_overhead = result["Overhead_Cost"].sum()
        total_cost = result["Total Cost"].sum()
        total_gm = result["GM"].sum()
        total_gm_pct = (
            total_gm / total_revenue) if total_revenue != 0 else (-1.0 if total_gm < 0 else 0.0)

        r1c1, r1c2, r1c3 = st.columns(3)
        r1c1.metric("Revenue", f"{total_revenue:,.2f}")
        r1c2.metric("Agent Cost", f"{total_agent:,.2f}")
        r1c3.metric("Overhead Cost", f"{total_overhead:,.2f}")
        r2c1, r2c2, r2c3 = st.columns(3)
        r2c1.metric("Total Cost", f"{total_cost:,.2f}")
        r2c2.metric("GM", f"{total_gm:,.2f}")
        r2c3.metric("GM %", f"{total_gm_pct*100:,.2f}%")

        if enable_whatif and whatif_result is not None and not whatif_result.empty:
            wf_revenue = whatif_result["Revenue"].sum()
            wf_total_cost = whatif_result["Total Cost"].sum()
            wf_gm = whatif_result["GM"].sum()
            wf_gm_pct = (
                wf_gm / wf_revenue) if wf_revenue != 0 else (-1.0 if wf_gm < 0 else 0.0)
            cmp = pd.DataFrame({
                "Metric": ["Revenue", "Total Cost", "GM", "GM %"],
                "Base": [total_revenue, total_cost, total_gm, total_gm_pct],
                "What-if": [wf_revenue, wf_total_cost, wf_gm, wf_gm_pct],
            })
            cmp["Delta"] = cmp["What-if"] - cmp["Base"]
            cmp_show = cmp.copy()
            for m in ["Revenue", "Total Cost", "GM"]:
                mask = cmp_show["Metric"] == m
                cmp_show.loc[mask, "Base"] = cmp_show.loc[mask,
                                                          "Base"].map(lambda x: f"{x:,.2f}")
                cmp_show.loc[mask, "What-if"] = cmp_show.loc[mask,
                                                             "What-if"].map(lambda x: f"{x:,.2f}")
                cmp_show.loc[mask, "Delta"] = cmp_show.loc[mask,
                                                           "Delta"].map(lambda x: f"{x:,.2f}")
            mask_pct = cmp_show["Metric"] == "GM %"
            cmp_show.loc[mask_pct, "Base"] = (
                cmp.loc[mask_pct, "Base"] * 100).round(2).astype(str) + "%"
            cmp_show.loc[mask_pct, "What-if"] = (
                cmp.loc[mask_pct, "What-if"] * 100).round(2).astype(str) + "%"
            cmp_show.loc[mask_pct, "Delta"] = (
                (cmp.loc[mask_pct, "Delta"] * 100).round(2)).astype(str) + " pp"
            st.markdown("**Scenario Comparison**")
            st.dataframe(cmp_show, use_container_width=True)

        st.divider()
        st.markdown("**Trend (Monthly)**")
        trend = result.groupby("Month", as_index=False)[
            ["Revenue", "Total Cost", "GM"]].sum().sort_values("Month")
        trend["GM_%"] = gm_pct_from_series(trend["GM"], trend["Revenue"])

        trend_base = trend.copy()
        trend_base["Scenario"] = "Base"
        trend_all = trend_base.copy()

        if enable_whatif and whatif_result is not None and not whatif_result.empty:
            wf_trend = whatif_result.groupby("Month", as_index=False)[
                ["Revenue", "Total Cost", "GM"]].sum().sort_values("Month")
            wf_trend["GM_%"] = gm_pct_from_series(
                wf_trend["GM"], wf_trend["Revenue"])
            wf_trend["Scenario"] = "What-if"
            trend_all = pd.concat([trend_base, wf_trend], ignore_index=True)

        trend_amt = trend_all.melt(
            id_vars=["Month", "Scenario"],
            value_vars=["Revenue", "Total Cost", "GM"],
            var_name="Metric",
            value_name="Value"
        )
        chart_amt = alt.Chart(trend_amt).mark_line(point=True).encode(
            x=alt.X("Month:N", title="Month"),
            y=alt.Y("Value:Q", title="TRY"),
            color=alt.Color("Metric:N", title="Metric"),
            strokeDash=alt.StrokeDash("Scenario:N", title="Scenario"),
            tooltip=[
                alt.Tooltip("Month:N"),
                alt.Tooltip("Scenario:N"),
                alt.Tooltip("Metric:N"),
                alt.Tooltip("Value:Q", format=",.2f"),
            ],
        ).properties(height=280)
        st.altair_chart(chart_amt, use_container_width=True)

        chart_gm = alt.Chart(trend_all).mark_line(point=True).encode(
            x=alt.X("Month:N", title="Month"),
            y=alt.Y("GM_%:Q", title="GM %", axis=alt.Axis(format=".0%")),
            color=alt.Color("Scenario:N", title="Scenario"),
            tooltip=[
                alt.Tooltip("Month:N"),
                alt.Tooltip("Scenario:N"),
                alt.Tooltip("GM_%:Q", format=".2%"),
            ],
        ).properties(height=220)
        st.altair_chart(chart_gm, use_container_width=True)

        st.divider()
        st.markdown("**Detailed Results**")
        show = result.copy()
        show["COLA_%"] = (show["COLA_%"] * 100).round(2).astype(str) + "%"
        show["GM_%"] = (show["GM_%"] * 100).round(2).astype(str) + "%"

        for col in [
            "Production_Hours", "FTE", "Agent_Cost", "Overhead_Cost", "Total Cost",
            "UnitPrice", "Adj. Unit Price", "FX_Rate", "Eff_Production_Hours", "Eff_FTE",
            "Eff_FTE_Hours", "Billable_Hours", "Revenue", "GM"
        ]:
            if col in show.columns:
                dec = 4 if col == "FX_Rate" else 2
                show[col] = money_fmt(show[col], decimals=dec)

        ordered = [
            "Month", "Account", "Language", "Production_Hours", "FTE",
            "Agent_Cost", "Overhead_Cost", "Total Cost",
            "UnitPrice", "COLA_%", "Adj. Unit Price", "Billing_Mode", "Currency", "FX_Rate",
            "Eff_Production_Hours", "Eff_FTE", "Eff_FTE_Hours", "Billable_Hours",
            "Revenue", "GM", "GM_%", "Notes"
        ]
        ordered = [c for c in ordered if c in show.columns]
        show = show[ordered]
        st.dataframe(show, use_container_width=True)

        st.markdown("**Summary by Account**")
        s1 = result.groupby("Account", as_index=False)[
            ["Revenue", "Total Cost", "GM"]].sum()
        s1["GM_%"] = gm_pct_from_series(s1["GM"], s1["Revenue"])
        s1_show = s1.copy()
        s1_show["GM_%"] = (s1_show["GM_%"] * 100).round(2).astype(str) + "%"
        for c in ["Revenue", "Total Cost", "GM"]:
            s1_show[c] = money_fmt(s1_show[c])
        st.dataframe(s1_show, use_container_width=True)

        st.markdown("**Summary by Account & Language**")
        s2 = result.groupby(["Account", "Language"], as_index=False)[
            ["Revenue", "Total Cost", "GM"]].sum()
        s2["GM_%"] = gm_pct_from_series(s2["GM"], s2["Revenue"])
        s2_show = s2.copy()
        s2_show["GM_%"] = (s2_show["GM_%"] * 100).round(2).astype(str) + "%"
        for c in ["Revenue", "Total Cost", "GM"]:
            s2_show[c] = money_fmt(s2_show[c])
        st.dataframe(s2_show, use_container_width=True)

        st.divider()
        # Capacity export tables (same logic as Capacity tab)
        cap_export_detail = pd.DataFrame()
        cap_export_summary = pd.DataFrame()
        cap_export = st.session_state.plan_df.copy()
        if not cap_export.empty:
            cap_export["Month"] = cap_export["Month"].astype(str)
            cap_export["FTE"] = pd.to_numeric(
                cap_export["FTE"], errors="coerce").fillna(0.0)
            cap_export["Production_Hours"] = pd.to_numeric(
                cap_export["Production_Hours"], errors="coerce").fillna(0.0)
            cap_export["Workdays"] = cap_export["Month"].apply(
                workdays_in_month)
            cap_export["Hours_Per_FTE"] = cap_export["Workdays"] * \
                cfg["hours_per_workday"]
            cap_export["FTE_Capacity_Hours"] = cap_export["FTE"] * \
                cap_export["Hours_Per_FTE"]
            cap_export["Eff_Capacity_Hours"] = cap_export["FTE_Capacity_Hours"] * \
                (1 - cfg["shrinkage_pct"])
            cap_export["Required_Hours"] = cap_export["Production_Hours"]
            cap_export["Gap_Hours"] = cap_export["Eff_Capacity_Hours"] - \
                cap_export["Required_Hours"]
            cap_export["Buffer_Hours"] = cap_export["Gap_Hours"].clip(lower=0)
            cap_export["Over_Demand_Hours"] = (
                cap_export["Required_Hours"] - cap_export["Eff_Capacity_Hours"]).clip(lower=0)
            cap_export["Idle_Hours"] = (
                cap_export["Eff_Capacity_Hours"] - cap_export["Required_Hours"]).clip(lower=0)
            cap_export["Invoiceable_Hours"] = np.minimum(
                cap_export["Required_Hours"], cap_export["Eff_Capacity_Hours"])
            cap_export["Buffer_%"] = np.where(
                cap_export["Eff_Capacity_Hours"] == 0, 0.0, cap_export["Buffer_Hours"] /
                cap_export["Eff_Capacity_Hours"]
            )
            cap_export["Status"] = np.where(
                cap_export["Over_Demand_Hours"] > 0,
                "Under Capacity (Cannot Invoice Full Demand)",
                "Over Capacity / Idle",
            )
            cap_budget_export = compute_budget(
                st.session_state.plan_df,
                st.session_state.unit_prices_df,
                st.session_state.cost_defaults_df,
                st.session_state.fx_rates_df,
                st.session_state.overhead_df,
                cfg,
            )
            if not cap_budget_export.empty:
                merge_cols = ["Month", "Account", "Language",
                              "Adj. Unit Price", "FX_Rate", "Total Cost"]
                merge_cols = [
                    c for c in merge_cols if c in cap_budget_export.columns]
                cap_export = cap_export.merge(
                    cap_budget_export[merge_cols],
                    on=["Month", "Account", "Language"],
                    how="left",
                )
            cap_export["Adj. Unit Price"] = pd.to_numeric(
                cap_export.get("Adj. Unit Price"), errors="coerce").fillna(0.0)
            cap_export["FX_Rate"] = pd.to_numeric(
                cap_export.get("FX_Rate"), errors="coerce").fillna(1.0)
            cap_export["Total Cost"] = pd.to_numeric(
                cap_export.get("Total Cost"), errors="coerce").fillna(0.0)
            cap_export["Rate_TRY_per_Hour"] = cap_export["Adj. Unit Price"] * \
                cap_export["FX_Rate"]
            cap_export["Potential_Revenue_TRY"] = cap_export["Required_Hours"] * \
                cap_export["Rate_TRY_per_Hour"]
            cap_export["Invoiceable_Revenue_TRY"] = cap_export["Invoiceable_Hours"] * \
                cap_export["Rate_TRY_per_Hour"]
            cap_export["Lost_Revenue_TRY"] = (
                cap_export["Potential_Revenue_TRY"] - cap_export["Invoiceable_Revenue_TRY"]).clip(lower=0)
            cap_export["Cost_per_Capacity_Hour"] = np.where(
                cap_export["Eff_Capacity_Hours"] == 0, 0.0, cap_export["Total Cost"] /
                cap_export["Eff_Capacity_Hours"]
            )
            cap_export["Idle_Cost_TRY"] = cap_export["Idle_Hours"] * \
                cap_export["Cost_per_Capacity_Hour"]
            cap_export["Actual_GM_TRY"] = cap_export["Invoiceable_Revenue_TRY"] - \
                cap_export["Total Cost"]
            cap_export["Potential_GM_TRY"] = cap_export["Potential_Revenue_TRY"] - \
                cap_export["Total Cost"]
            cap_export["Actual_GM_%"] = np.where(
                cap_export["Invoiceable_Revenue_TRY"] != 0,
                cap_export["Actual_GM_TRY"] /
                cap_export["Invoiceable_Revenue_TRY"],
                np.where(cap_export["Actual_GM_TRY"] < 0, -1.0, 0.0),
            )
            cap_export["Potential_GM_%"] = np.where(
                cap_export["Potential_Revenue_TRY"] != 0,
                cap_export["Potential_GM_TRY"] /
                cap_export["Potential_Revenue_TRY"],
                np.where(cap_export["Potential_GM_TRY"] < 0, -1.0, 0.0),
            )
            cap_export_detail = cap_export.copy()
            cap_export_summary = cap_export.groupby(["Month", "Account"], as_index=False)[
                [
                    "Eff_Capacity_Hours", "Required_Hours", "Invoiceable_Hours", "Over_Demand_Hours",
                    "Idle_Hours", "Gap_Hours", "Buffer_Hours", "Potential_Revenue_TRY",
                    "Invoiceable_Revenue_TRY", "Lost_Revenue_TRY", "Idle_Cost_TRY", "Total Cost",
                    "Actual_GM_TRY", "Potential_GM_TRY",
                ]
            ].sum()
            cap_export_summary["Actual_GM_%"] = np.where(
                cap_export_summary["Invoiceable_Revenue_TRY"] != 0,
                cap_export_summary["Actual_GM_TRY"] /
                cap_export_summary["Invoiceable_Revenue_TRY"],
                np.where(cap_export_summary["Actual_GM_TRY"] < 0, -1.0, 0.0),
            )
            cap_export_summary["Potential_GM_%"] = np.where(
                cap_export_summary["Potential_Revenue_TRY"] != 0,
                cap_export_summary["Potential_GM_TRY"] /
                cap_export_summary["Potential_Revenue_TRY"],
                np.where(
                    cap_export_summary["Potential_GM_TRY"] < 0, -1.0, 0.0),
            )

        export_bytes = df_to_excel_bytes(
            {
                "Input_UnitPrices": st.session_state.unit_prices_df,
                "Input_CostDefaults": st.session_state.cost_defaults_df,
                "Input_Overhead": st.session_state.overhead_df,
                "Input_Plan": st.session_state.plan_df,
                "Input_FX": st.session_state.fx_rates_df if "fx_rates_df" in st.session_state else pd.DataFrame(columns=["Month", "Currency", "FX_Rate"]),
                "Detailed": result,
                "Summary_Account": s1,
                "Summary_Acc_Lang": s2,
                "Trend": trend,
                "Capacity_Detail": cap_export_detail,
                "Capacity_Summary": cap_export_summary,
            }
        )
        st.download_button(
            "⬇️ Download Results (xlsx)",
            data=export_bytes,
            file_name="budget_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

with tab_capacity:
    st.subheader("FTE Capacity vs Operational Hours")
    st.caption(
        f"Capacity is calculated as `FTE × {cfg['hours_per_workday']:.2f} hours × workdays (Mon-Fri)` per month.")

    cap_df = st.session_state.plan_df.copy()
    if cap_df.empty:
        st.info("No Monthly Plan data yet.")
    else:
        cap_df["Month"] = cap_df["Month"].astype(str)
        cap_df["FTE"] = pd.to_numeric(
            cap_df["FTE"], errors="coerce").fillna(0.0)
        cap_df["Production_Hours"] = pd.to_numeric(
            cap_df["Production_Hours"], errors="coerce").fillna(0.0)
        cap_df["Workdays"] = cap_df["Month"].apply(workdays_in_month)
        cap_df["Hours_Per_FTE"] = cap_df["Workdays"] * cfg["hours_per_workday"]
        cap_df["FTE_Capacity_Hours"] = cap_df["FTE"] * cap_df["Hours_Per_FTE"]
        cap_df["Eff_Capacity_Hours"] = cap_df["FTE_Capacity_Hours"] * \
            (1 - cfg["shrinkage_pct"])
        cap_df["Required_Hours"] = cap_df["Production_Hours"]
        cap_df["Gap_Hours"] = cap_df["Eff_Capacity_Hours"] - \
            cap_df["Required_Hours"]
        cap_df["Buffer_Hours"] = cap_df["Gap_Hours"].clip(lower=0)
        cap_df["Over_Demand_Hours"] = (
            cap_df["Required_Hours"] - cap_df["Eff_Capacity_Hours"]).clip(lower=0)
        cap_df["Idle_Hours"] = (
            cap_df["Eff_Capacity_Hours"] - cap_df["Required_Hours"]).clip(lower=0)
        cap_df["Invoiceable_Hours"] = np.minimum(
            cap_df["Required_Hours"], cap_df["Eff_Capacity_Hours"])
        cap_df["Buffer_%"] = np.where(
            cap_df["Eff_Capacity_Hours"] == 0,
            0.0,
            cap_df["Buffer_Hours"] / cap_df["Eff_Capacity_Hours"],
        )
        cap_df["Status"] = np.where(
            cap_df["Over_Demand_Hours"] > 0,
            "Under Capacity (Cannot Invoice Full Demand)",
            "Over Capacity / Idle",
        )

        # Financial impact view (using rate/cost context from budget result)
        cap_budget = compute_budget(
            st.session_state.plan_df,
            st.session_state.unit_prices_df,
            st.session_state.cost_defaults_df,
            st.session_state.fx_rates_df,
            st.session_state.overhead_df,
            cfg,
        )
        if not cap_budget.empty:
            merge_cols = [
                "Month", "Account", "Language",
                "Adj. Unit Price", "FX_Rate", "Total Cost",
                "Billing_Mode", "Revenue", "GM", "Billable_Hours",
            ]
            merge_cols = [c for c in merge_cols if c in cap_budget.columns]
            cap_df = cap_df.merge(
                cap_budget[merge_cols],
                on=["Month", "Account", "Language"],
                how="left",
            )
        else:
            cap_df["Adj. Unit Price"] = 0.0
            cap_df["FX_Rate"] = 1.0
            cap_df["Total Cost"] = 0.0
            cap_df["Billing_Mode"] = "Unit Price × Production Hours"
            cap_df["Revenue"] = 0.0
            cap_df["GM"] = 0.0
            cap_df["Billable_Hours"] = 0.0

        cap_df["Adj. Unit Price"] = pd.to_numeric(
            cap_df.get("Adj. Unit Price"), errors="coerce").fillna(0.0)
        cap_df["FX_Rate"] = pd.to_numeric(
            cap_df.get("FX_Rate"), errors="coerce").fillna(1.0)
        cap_df["Total Cost"] = pd.to_numeric(
            cap_df.get("Total Cost"), errors="coerce").fillna(0.0)
        cap_df["Revenue"] = pd.to_numeric(
            cap_df.get("Revenue"), errors="coerce").fillna(0.0)
        cap_df["Billable_Hours"] = pd.to_numeric(
            cap_df.get("Billable_Hours"), errors="coerce").fillna(0.0)

        cap_df["Rate_TRY_per_Hour"] = np.where(
            cap_df["Billable_Hours"] == 0,
            cap_df["Adj. Unit Price"] * cap_df["FX_Rate"],
            cap_df["Revenue"] / cap_df["Billable_Hours"],
        )
        is_prod_mode = cap_df["Billing_Mode"] != "Unit Price × FTE"
        cap_df["Potential_Revenue_TRY"] = np.where(
            is_prod_mode,
            cap_df["Required_Hours"] * cap_df["Rate_TRY_per_Hour"],
            cap_df["Revenue"],
        )
        cap_df["Invoiceable_Revenue_TRY"] = np.where(
            is_prod_mode,
            cap_df["Invoiceable_Hours"] * cap_df["Rate_TRY_per_Hour"],
            cap_df["Revenue"],
        )
        cap_df["Lost_Revenue_TRY"] = np.where(
            is_prod_mode,
            (cap_df["Potential_Revenue_TRY"] -
             cap_df["Invoiceable_Revenue_TRY"]).clip(lower=0),
            0.0,
        )
        cap_df["Cost_per_Capacity_Hour"] = np.where(
            cap_df["Eff_Capacity_Hours"] == 0, 0.0, cap_df["Total Cost"] /
            cap_df["Eff_Capacity_Hours"]
        )
        cap_df["Idle_Cost_TRY"] = cap_df["Idle_Hours"] * \
            cap_df["Cost_per_Capacity_Hour"]
        cap_df["Actual_GM_TRY"] = cap_df["Revenue"] - cap_df["Total Cost"]
        cap_df["Potential_GM_TRY"] = cap_df["Potential_Revenue_TRY"] - \
            cap_df["Total Cost"]
        cap_df["Actual_GM_%"] = gm_pct_from_series(
            cap_df["Actual_GM_TRY"], cap_df["Revenue"])
        cap_df["Potential_GM_%"] = np.where(
            cap_df["Potential_Revenue_TRY"] != 0,
            cap_df["Potential_GM_TRY"] / cap_df["Potential_Revenue_TRY"],
            np.where(cap_df["Potential_GM_TRY"] < 0, -1.0, 0.0),
        )

        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Total Over-Demand Hours",
                  f"{cap_df['Over_Demand_Hours'].sum():,.2f}")
        k2.metric("Total Idle Hours", f"{cap_df['Idle_Hours'].sum():,.2f}")
        k3.metric("Lost Invoice Revenue (TRY)",
                  f"{cap_df['Lost_Revenue_TRY'].sum():,.2f}")
        total_actual_rev = cap_df["Revenue"].sum()
        total_actual_gm = cap_df["Actual_GM_TRY"].sum()
        total_actual_gm_pct = (
            total_actual_gm / total_actual_rev
        ) if total_actual_rev != 0 else (-1.0 if total_actual_gm < 0 else 0.0)
        k4.metric("Actual GM %", f"{total_actual_gm_pct*100:,.2f}%")
        total_potential_rev = cap_df["Potential_Revenue_TRY"].sum()
        total_potential_gm = cap_df["Potential_GM_TRY"].sum()
        total_potential_gm_pct = (
            total_potential_gm / total_potential_rev
        ) if total_potential_rev != 0 else (-1.0 if total_potential_gm < 0 else 0.0)
        k5.metric("Potential GM % (No Lost Hours)",
                  f"{total_potential_gm_pct*100:,.2f}%")

        st.markdown("**Detailed Capacity**")
        cap_show = cap_df.copy()
        for c in [
            "Hours_Per_FTE", "FTE_Capacity_Hours", "Eff_Capacity_Hours", "Required_Hours",
            "Gap_Hours", "Buffer_Hours", "Over_Demand_Hours", "Idle_Hours",
            "Invoiceable_Hours", "Potential_Revenue_TRY", "Invoiceable_Revenue_TRY",
            "Lost_Revenue_TRY", "Idle_Cost_TRY", "Actual_GM_TRY", "Potential_GM_TRY"
        ]:
            cap_show[c] = money_fmt(cap_show[c], decimals=2)
        cap_show["Buffer_%"] = (pd.to_numeric(cap_df["Buffer_%"], errors="coerce").fillna(
            0.0) * 100).round(2).astype(str) + "%"
        cap_show["Actual_GM_%"] = (pd.to_numeric(cap_df["Actual_GM_%"], errors="coerce").fillna(
            0.0) * 100).round(2).astype(str) + "%"
        cap_show["Potential_GM_%"] = (pd.to_numeric(cap_df["Potential_GM_%"], errors="coerce").fillna(
            0.0) * 100).round(2).astype(str) + "%"
        detail_cols = [
            "Month", "Account", "Language", "FTE", "Workdays", "Hours_Per_FTE",
            "Eff_Capacity_Hours", "Required_Hours", "Invoiceable_Hours",
            "Over_Demand_Hours", "Idle_Hours", "Gap_Hours", "Buffer_Hours", "Buffer_%",
            "Potential_Revenue_TRY", "Invoiceable_Revenue_TRY", "Lost_Revenue_TRY", "Idle_Cost_TRY",
            "Actual_GM_%", "Potential_GM_%", "Status"
        ]
        st.dataframe(cap_show[detail_cols], use_container_width=True)

        st.markdown("**Summary by Month + Operation**")
        cap_sum = cap_df.groupby(["Month", "Account"], as_index=False)[
            [
                "Eff_Capacity_Hours", "Required_Hours", "Invoiceable_Hours",
                "Over_Demand_Hours", "Idle_Hours", "Gap_Hours", "Buffer_Hours",
                "Potential_Revenue_TRY", "Invoiceable_Revenue_TRY", "Lost_Revenue_TRY", "Idle_Cost_TRY", "Total Cost", "Revenue"
            ]
        ].sum()
        cap_sum["Actual_GM_TRY"] = cap_sum["Revenue"] - cap_sum["Total Cost"]
        cap_sum["Potential_GM_TRY"] = cap_sum["Potential_Revenue_TRY"] - \
            cap_sum["Total Cost"]
        cap_sum["Actual_GM_%"] = gm_pct_from_series(
            cap_sum["Actual_GM_TRY"], cap_sum["Revenue"])
        cap_sum["Potential_GM_%"] = np.where(
            cap_sum["Potential_Revenue_TRY"] != 0,
            cap_sum["Potential_GM_TRY"] / cap_sum["Potential_Revenue_TRY"],
            np.where(cap_sum["Potential_GM_TRY"] < 0, -1.0, 0.0),
        )
        cap_sum["Buffer_%"] = np.where(
            cap_sum["Eff_Capacity_Hours"] == 0, 0.0, cap_sum["Buffer_Hours"] /
            cap_sum["Eff_Capacity_Hours"]
        )
        cap_sum["Status"] = np.where(
            cap_sum["Over_Demand_Hours"] > 0,
            "Under Capacity (Cannot Invoice Full Demand)",
            "Over Capacity / Idle",
        )
        cap_sum_show = cap_sum.copy()
        for c in [
            "Eff_Capacity_Hours", "Required_Hours", "Invoiceable_Hours", "Over_Demand_Hours",
            "Idle_Hours", "Gap_Hours", "Buffer_Hours", "Potential_Revenue_TRY",
            "Invoiceable_Revenue_TRY", "Lost_Revenue_TRY", "Idle_Cost_TRY", "Total Cost",
            "Actual_GM_TRY", "Potential_GM_TRY"
        ]:
            cap_sum_show[c] = money_fmt(cap_sum_show[c], decimals=2)
        cap_sum_show["Buffer_%"] = (
            cap_sum["Buffer_%"] * 100).round(2).astype(str) + "%"
        cap_sum_show["Actual_GM_%"] = (
            cap_sum["Actual_GM_%"] * 100).round(2).astype(str) + "%"
        cap_sum_show["Potential_GM_%"] = (
            cap_sum["Potential_GM_%"] * 100).round(2).astype(str) + "%"
        st.dataframe(cap_sum_show, use_container_width=True)

        st.markdown("**Capacity Gap Trend (Hours)**")
        cap_chart = cap_sum.groupby("Month", as_index=False)["Gap_Hours"].sum()
        gap_chart = alt.Chart(cap_chart).mark_bar().encode(
            x=alt.X("Month:N", title="Month"),
            y=alt.Y("Gap_Hours:Q", title="Gap Hours (Capacity - Required)"),
            color=alt.condition(alt.datum.Gap_Hours < 0, alt.value(
                "#d62728"), alt.value("#2ca02c")),
            tooltip=[
                alt.Tooltip("Month:N"),
                alt.Tooltip("Gap_Hours:Q", format=",.2f"),
            ],
        ).properties(height=240)
        st.altair_chart(gap_chart, use_container_width=True)
