import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="Ops Budget", layout="wide")

# -----------------------------
# Helpers
# -----------------------------


def month_range(start_ym: str, months: int) -> pd.DatetimeIndex:
    """start_ym like '2026-02'"""
    start = pd.to_datetime(start_ym + "-01")
    return pd.date_range(start=start, periods=months, freq="MS")


def make_unique_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Ensure unique column names by appending .1, .2 where needed.
    """
    cols = list(df.columns)
    seen = {}
    new_cols = []
    for c in cols:
        if c not in seen:
            seen[c] = 0
            new_cols.append(c)
        else:
            seen[c] += 1
            new_cols.append(f"{c}.{seen[c]}")
    df = df.copy()
    df.columns = new_cols
    return df


def df_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Sheet1") -> bytes:
    output = BytesIO()
    try:
        engine = "xlsxwriter"
        __import__("xlsxwriter")
    except Exception:
        engine = "openpyxl"
    with pd.ExcelWriter(output, engine=engine) as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()


def format_number_cols(df: pd.DataFrame, cols: list, decimals: int = 2) -> pd.DataFrame:
    out = df.copy()
    for c in cols:
        if c in out.columns:
            out[c] = pd.to_numeric(out[c], errors="coerce").fillna(0.0).map(
                lambda x: f"{x:,.{decimals}f}"
            )
    return out


def compute_budget(prices_df, plan_df, cola_cfg, fx_rates_df):
    df = plan_df.copy()

    # COLA applies after selected month
    df["Month_dt"] = pd.to_datetime(df["Month"] + "-01", errors="coerce")
    cola_start_dt = pd.to_datetime(
        cola_cfg["cola_start_month"] + "-01", errors="coerce")
    if pd.isna(cola_start_dt):
        cola_start_dt = df["Month_dt"].min()
    df["COLA_%"] = np.where(
        df["Month_dt"] >= cola_start_dt, cola_cfg["cola_pct"], 0.0)

    if "Overhead_Cost" not in df.columns:
        df["Overhead_Cost"] = 0.0
    df["Adj_Cost"] = df["Base_Cost"] + df["Overhead_Cost"]

    # Revenue from Hours * Unit Price (by Account & Language)
    # Expand plan rows by matching Account/Language price
    df = df.merge(
        prices_df.rename(columns={"UnitPrice": "Unit_Price"}),
        on=["Account", "Language"],
        how="left"
    )
    df["Unit_Price"] = df["Unit_Price"].fillna(0)
    df["Billing_Mode"] = df["Billing_Mode"].fillna(
        "Unit Price × Production Hours")
    # Convert Unit Price to TRY with monthly FX rates (if provided)
    fx = fx_rates_df.copy() if fx_rates_df is not None else pd.DataFrame(
        columns=["Month", "FX_Rate"])
    if "Month" in fx.columns and "FX_Rate" in fx.columns:
        fx["Month"] = fx["Month"].astype(str)
        df = df.merge(fx, on="Month", how="left")
    else:
        df["FX_Rate"] = 1.0
    df["FX_Rate"] = pd.to_numeric(
        df.get("FX_Rate", 1.0), errors="coerce").fillna(1.0)
    df["Unit_Price_TRY"] = df["Unit_Price"] * df["FX_Rate"]

    # Adj Unit Price is based on Unit Price and COLA only
    df["Adj_Unit_Price"] = df["Unit_Price"] * (1 + df["COLA_%"])

    # Volume can be hours
    df["Eff_Production_Hours"] = df["Production_Hours"] * \
        (1 - cola_cfg["shrinkage_pct"])
    df["Eff_FTE"] = df["FTE"] * (1 - cola_cfg["shrinkage_pct"])
    df["Eff_FTE_Hours"] = df["Eff_FTE"] * 180.0

    df["Billable_Hours"] = np.where(
        df["Billing_Mode"] == "Unit Price × FTE",
        df["Eff_FTE_Hours"],
        df["Eff_Production_Hours"]
    )
    # Revenue in TRY using FX rate
    df["Revenue"] = df["Billable_Hours"] * df["Adj_Unit_Price"] * df["FX_Rate"]
    df["GM"] = df["Revenue"] - df["Adj_Cost"]
    df["GM_%"] = np.where(df["Revenue"] == 0, 0, df["GM"] / df["Revenue"])

    return df


# -----------------------------
# Defaults / Session State
# -----------------------------
ACCOUNTS = ["Sky X Chat", "SALT IB", "Inditex",
            "BackMarket", "EMMA", "Chrono24", "Adidas", "TP Vision"]
LANGUAGES = ["DE", "FR", "IT", "TR", "EN", "ES", "NL"]

if "prices_df" not in st.session_state:
    st.session_state.prices_df = pd.DataFrame(
        [{"Account": a, "Language": l, "UnitPrice": 0.0, "Billing_Mode": "Unit Price × Production Hours"}
            for a in ACCOUNTS for l in ["DE", "FR", "EN"]]
    )

if "plan_df" not in st.session_state:
    months = month_range("2026-02", 12)
    st.session_state.plan_df = pd.DataFrame({
        "Month": months.strftime("%Y-%m"),
        "Account": ["Sky X Chat"]*len(months),
        "Language": ["DE"]*len(months),
        "Production_Hours": [0.0]*len(months),
        "FTE": [0.0]*len(months),
        "Base_Cost": [0.0]*len(months),  # cost for that month row
        "Overhead_Cost": [0.0]*len(months),
        "Notes": [""]*len(months)
    })

if "base_cost_df" not in st.session_state:
    st.session_state.base_cost_df = pd.DataFrame({
        "Language": ["DE"],
        "FTE": [1.0],
        "Salary": [0.0],
        "OSS": [2083.0],
        "Food": [5850.0],
        "Goalpex_%": [0.10],
        "Additional_Cost": [0.0],
    })

if "overhead_cost_df" not in st.session_state:
    st.session_state.overhead_cost_df = pd.DataFrame({
        "Account": ["Sky X Chat"],
        "Role": ["Operation Manager"],
        "FTE": [1.0],
        "Salary": [0.0],
        "OSS": [2083.0],
        "Food": [5850.0],
        "Goalpex_%": [0.10],
        "Additional_Cost": [0.0],
    })

if "fx_rates_df" not in st.session_state:
    if "plan_df" in st.session_state:
        months = sorted(st.session_state.plan_df["Month"].astype(
            str).unique().tolist())
        st.session_state.fx_rates_df = pd.DataFrame({
            "Month": months,
            "FX_Rate": [1.0]*len(months)
        })
    else:
        st.session_state.fx_rates_df = pd.DataFrame(
            {"Month": [], "FX_Rate": []})

if "config_defaults" not in st.session_state:
    st.session_state.config_defaults = {
        "OSS": 2083.0,
        "Food": 5850.0,
        "Goalpex_%": 0.10,
    }

if "lang_defaults_df" not in st.session_state:
    st.session_state.lang_defaults_df = pd.DataFrame({
        "Language": ["DE", "FR", "EN"],
        "Salary": [0.0, 0.0, 0.0],
        "OSS": [2083.0, 2083.0, 2083.0],
        "Food": [5850.0, 5850.0, 5850.0],
        "Goalpex_%": [0.10, 0.10, 0.10],
    })

if "role_defaults_df" not in st.session_state:
    st.session_state.role_defaults_df = pd.DataFrame({
        "Role": ["Operation Manager", "Teamleader", "Trainer & Quality", "RTA", "Planner", "WFM", "Operation Support"],
        "Salary": [0.0]*7,
        "OSS": [2083.0]*7,
        "Food": [5850.0]*7,
        "Goalpex_%": [0.10]*7,
    })

# -----------------------------
# UI
# -----------------------------
st.title("Operations Budget Calculator")

with st.sidebar:
    st.header("⚙️ Setup")

    st.subheader("Operation Filter")
    operation_filter = st.selectbox(
        "Show data for",
        options=["All"] + ACCOUNTS,
        index=0
    )

    start_ym = st.text_input("Start Month (YYYY-MM)", "2026-02")
    horizon = st.number_input(
        "Number of months", min_value=1, max_value=60, value=12)

    def rebuild_plan_for_operation(op_name: str, months_idx: pd.DatetimeIndex):
        base = st.session_state.plan_df.copy()
        base = base[base["Account"] != op_name]
        new_rows = pd.DataFrame({
            "Month": months_idx.strftime("%Y-%m"),
            "Account": [op_name]*len(months_idx),
            "Language": ["DE"]*len(months_idx),
            "Production_Hours": [0.0]*len(months_idx),
            "FTE": [0.0]*len(months_idx),
            "Base_Cost": [0.0]*len(months_idx),
            "Overhead_Cost": [0.0]*len(months_idx),
            "Notes": [""]*len(months_idx)
        })
        st.session_state.plan_df = pd.concat(
            [base, new_rows], ignore_index=True)

    # Sync months without wiping existing values
    months = month_range(start_ym, int(horizon))
    months_list = months.strftime("%Y-%m").tolist()

    def ensure_months_for_operation(op_name: str):
        df = st.session_state.plan_df.copy()
        existing = df[df["Account"] == op_name]
        existing_months = set(existing["Month"].astype(str))
        missing = [m for m in months_list if m not in existing_months]
        if missing:
            new_rows = pd.DataFrame({
                "Month": missing,
                "Account": [op_name]*len(missing),
                "Language": ["DE"]*len(missing),
                "Production_Hours": [0.0]*len(missing),
                "FTE": [0.0]*len(missing),
                "Base_Cost": [0.0]*len(missing),
                "Overhead_Cost": [0.0]*len(missing),
                "Notes": [""]*len(missing)
            })
            st.session_state.plan_df = pd.concat(
                [df, new_rows], ignore_index=True)

        # Optionally trim months outside range for this operation
        df = st.session_state.plan_df.copy()
        df_op = df[df["Account"] == op_name]
        df_keep = df[~((df["Account"] == op_name) & (
            ~df["Month"].astype(str).isin(months_list)))]
        st.session_state.plan_df = df_keep

    if operation_filter != "All":
        ensure_months_for_operation(operation_filter)
    else:
        for op in ACCOUNTS:
            ensure_months_for_operation(op)

    if st.button("🔄 Reset months (rebuild tables)"):
        if operation_filter != "All":
            rebuild_plan_for_operation(operation_filter, months)
        else:
            for op in ACCOUNTS:
                rebuild_plan_for_operation(op, months)

    st.divider()
    st.subheader("COLA %")
    cola_pct = st.number_input(
        "COLA % (decimal)", value=0.0, step=0.01, format="%.4f")

    # Apply COLA after selected month
    months_for_cola = sorted(
        st.session_state.plan_df["Month"].unique().tolist())
    cola_start_month = st.selectbox(
        "Apply COLA starting month", options=months_for_cola, index=0)

    st.divider()
    st.subheader("Brut Multiplier")
    brut_multiplier = st.slider("Brut Multiplier", 1.0, 3.0, 1.58, 0.01)

    st.divider()
    st.subheader("Shrinkage %")
    shrinkage_pct = st.slider("Shrinkage %", 0.0, 100.0, 0.0, 0.5) / 100.0

    cola_cfg = {
        "cola_pct": cola_pct,
        "cola_start_month": cola_start_month,
        "fx_currency": "TRY",
        "fx_rate": 1.0,
        "shrinkage_pct": shrinkage_pct,
        "brut_multiplier": brut_multiplier
    }

    st.divider()
    st.subheader("Export")
    st.caption("After calculation, use the download buttons.")

tab0, tab1, tab2, tab3, tab4 = st.tabs(
    ["0) Config", "1) Unit Prices", "2) Base Cost", "3) Monthly Plan", "4) Results"])

# -----------------------------
# 0) Config
# -----------------------------
with tab0:
    st.subheader("Defaults & Configuration")
    st.caption(
        "Adjust default OSS, Food, Goalpex% and default salaries by Language and Role.")

    with st.form("config_form"):
        c1, c2, c3 = st.columns(3)
        with c1:
            default_oss = st.number_input("Default OSS (TRY)", value=float(
                st.session_state.config_defaults["OSS"]), step=10.0, format="%.2f")
        with c2:
            default_food = st.number_input("Default Food (TRY)", value=float(
                st.session_state.config_defaults["Food"]), step=10.0, format="%.2f")
        with c3:
            default_goalpex = st.number_input("Default Goalpex %", value=float(
                st.session_state.config_defaults["Goalpex_%"]), step=0.01, format="%.2f")

        st.divider()
        st.subheader("Language Defaults")
        lang_defaults = st.data_editor(
            st.session_state.lang_defaults_df,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "Language": st.column_config.SelectboxColumn("Language", options=LANGUAGES),
                "Salary": st.column_config.NumberColumn("Salary (TRY)", min_value=0.0, step=100.0, format="%.2f"),
                "OSS": st.column_config.NumberColumn("OSS (TRY)", min_value=0.0, step=10.0, format="%.2f"),
                "Food": st.column_config.NumberColumn("Food (TRY)", min_value=0.0, step=10.0, format="%.2f"),
                "Goalpex_%": st.column_config.NumberColumn("Goalpex %", min_value=0.0, step=0.01, format="%.2f"),
            }
        )

        st.subheader("Role Defaults")
        role_defaults = st.data_editor(
            st.session_state.role_defaults_df,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "Role": st.column_config.TextColumn("Role"),
                "Salary": st.column_config.NumberColumn("Salary (TRY)", min_value=0.0, step=100.0, format="%.2f"),
                "OSS": st.column_config.NumberColumn("OSS (TRY)", min_value=0.0, step=10.0, format="%.2f"),
                "Food": st.column_config.NumberColumn("Food (TRY)", min_value=0.0, step=10.0, format="%.2f"),
                "Goalpex_%": st.column_config.NumberColumn("Goalpex %", min_value=0.0, step=0.01, format="%.2f"),
            }
        )

        apply_defaults_to_tables = st.checkbox(
            "Apply defaults to Language/Role tables (fill blanks/zeros)", value=True)
        overwrite_defaults_tables = st.checkbox(
            "Force overwrite OSS/Food/Goalpex in tables", value=False)
        save_config = st.form_submit_button("Save Config")

    if save_config:
        st.session_state.config_defaults = {
            "OSS": default_oss,
            "Food": default_food,
            "Goalpex_%": default_goalpex
        }
        if apply_defaults_to_tables:
            for col, val in [("OSS", default_oss), ("Food", default_food), ("Goalpex_%", default_goalpex)]:
                if col in lang_defaults.columns:
                    lang_defaults[col] = pd.to_numeric(
                        lang_defaults[col], errors="coerce")
                    if overwrite_defaults_tables:
                        lang_defaults[col] = val
                    else:
                        lang_defaults[col] = lang_defaults[col].replace(
                            0, np.nan).fillna(val)
                if col in role_defaults.columns:
                    role_defaults[col] = pd.to_numeric(
                        role_defaults[col], errors="coerce")
                    if overwrite_defaults_tables:
                        role_defaults[col] = val
                    else:
                        role_defaults[col] = role_defaults[col].replace(
                            0, np.nan).fillna(val)
        st.session_state.lang_defaults_df = lang_defaults
        st.session_state.role_defaults_df = role_defaults
        st.success("Defaults updated.")

    st.divider()
    st.subheader("Bulk Apply Defaults")
    overwrite_existing = st.checkbox("Overwrite existing values", value=False)
    if st.button("Apply Config to existing rows"):
        # Apply to Base Cost rows by Language
        base_df = st.session_state.base_cost_df.copy()
        lang_defaults = st.session_state.lang_defaults_df.copy()
        lang_defaults["Language"] = lang_defaults["Language"].astype(
            str).str.strip()
        if not base_df.empty:
            base_df["Language"] = base_df["Language"].astype(str).str.strip()
            lang_salary = dict(
                zip(lang_defaults["Language"], lang_defaults["Salary"]))
            lang_oss = dict(
                zip(lang_defaults["Language"], lang_defaults["OSS"]))
            lang_food = dict(
                zip(lang_defaults["Language"], lang_defaults["Food"]))
            lang_goalpex = dict(
                zip(lang_defaults["Language"], lang_defaults["Goalpex_%"]))

            if overwrite_existing:
                base_df["Salary"] = base_df["Language"].map(
                    lang_salary).fillna(0.0)
                base_df["OSS"] = base_df["Language"].map(lang_oss).fillna(
                    st.session_state.config_defaults["OSS"])
                base_df["Food"] = base_df["Language"].map(lang_food).fillna(
                    st.session_state.config_defaults["Food"])
                base_df["Goalpex_%"] = base_df["Language"].map(lang_goalpex).fillna(
                    st.session_state.config_defaults["Goalpex_%"])
            else:
                base_df["Salary"] = base_df["Salary"].fillna(
                    base_df["Language"].map(lang_salary)).fillna(0.0)
                base_df["OSS"] = base_df["OSS"].fillna(base_df["Language"].map(
                    lang_oss)).fillna(st.session_state.config_defaults["OSS"])
                base_df["Food"] = base_df["Food"].fillna(base_df["Language"].map(
                    lang_food)).fillna(st.session_state.config_defaults["Food"])
                base_df["Goalpex_%"] = base_df["Goalpex_%"].fillna(base_df["Language"].map(
                    lang_goalpex)).fillna(st.session_state.config_defaults["Goalpex_%"])
            st.session_state.base_cost_df = base_df

        # Apply to Overhead rows by Role
        overhead_df = st.session_state.overhead_cost_df.copy()
        role_defaults = st.session_state.role_defaults_df.copy()
        role_defaults["Role"] = role_defaults["Role"].astype(str).str.strip()
        if not overhead_df.empty:
            overhead_df["Role"] = overhead_df["Role"].astype(str).str.strip()
            role_salary = dict(
                zip(role_defaults["Role"], role_defaults["Salary"]))
            role_oss = dict(zip(role_defaults["Role"], role_defaults["OSS"]))
            role_food = dict(zip(role_defaults["Role"], role_defaults["Food"]))
            role_goalpex = dict(
                zip(role_defaults["Role"], role_defaults["Goalpex_%"]))

            if overwrite_existing:
                overhead_df["Salary"] = overhead_df["Role"].map(
                    role_salary).fillna(0.0)
                overhead_df["OSS"] = overhead_df["Role"].map(
                    role_oss).fillna(st.session_state.config_defaults["OSS"])
                overhead_df["Food"] = overhead_df["Role"].map(
                    role_food).fillna(st.session_state.config_defaults["Food"])
                overhead_df["Goalpex_%"] = overhead_df["Role"].map(
                    role_goalpex).fillna(st.session_state.config_defaults["Goalpex_%"])
            else:
                overhead_df["Salary"] = overhead_df["Salary"].fillna(
                    overhead_df["Role"].map(role_salary)).fillna(0.0)
                overhead_df["OSS"] = overhead_df["OSS"].fillna(overhead_df["Role"].map(
                    role_oss)).fillna(st.session_state.config_defaults["OSS"])
                overhead_df["Food"] = overhead_df["Food"].fillna(overhead_df["Role"].map(
                    role_food)).fillna(st.session_state.config_defaults["Food"])
                overhead_df["Goalpex_%"] = overhead_df["Goalpex_%"].fillna(overhead_df["Role"].map(
                    role_goalpex)).fillna(st.session_state.config_defaults["Goalpex_%"])
            st.session_state.overhead_cost_df = overhead_df

        st.success("Applied config defaults to existing rows.")

# -----------------------------
# 1) Unit Prices
# -----------------------------
with tab1:
    st.subheader("Unit Price per Account & Language")

    st.info(
        "Define your Unit Price per language for each account (e.g., €/hour, €/FTE). "
        "You control what 'Volume' means in the plan tab."
    )

    prices_view = st.session_state.prices_df
    if operation_filter != "All":
        prices_view = prices_view[prices_view["Account"] == operation_filter]

    with st.form("unit_price_form"):
        prices_df = st.data_editor(
            prices_view,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "Account": st.column_config.SelectboxColumn("Account", options=ACCOUNTS),
                "Language": st.column_config.TextColumn("Language"),
                "UnitPrice": st.column_config.NumberColumn("Unit Price", min_value=0.0, step=0.01, format="%.2f"),
                "Billing_Mode": st.column_config.SelectboxColumn(
                    "Billing Mode",
                    options=["Unit Price × Production Hours",
                             "Unit Price × FTE"]
                ),
            }
        )
        save_prices = st.form_submit_button("Save Unit Prices")
    if save_prices:
        if "UnitPrice" in prices_df.columns:
            prices_df["UnitPrice"] = pd.to_numeric(
                prices_df["UnitPrice"], errors="coerce").fillna(0.0)
        # Merge edited view back into full table
        if operation_filter != "All":
            full = st.session_state.prices_df.copy()
            full = full[full["Account"] != operation_filter]
            st.session_state.prices_df = pd.concat(
                [full, prices_df], ignore_index=True)
        else:
            st.session_state.prices_df = prices_df

    st.divider()
    st.subheader("Unit Price FX (Monthly)")
    st.caption(
        "Enter monthly FX rates to convert Unit Price to TRY (used in Results).")
    fx_currency = st.selectbox("Unit Price Currency", [
                               "TRY", "EUR", "USD"], index=0, key="fx_currency_unitprice")
    fx_rates_df = st.session_state.fx_rates_df.copy()
    if "Month" not in fx_rates_df.columns or "FX_Rate" not in fx_rates_df.columns:
        fx_rates_df = pd.DataFrame({"Month": [], "FX_Rate": []})
    # ensure months present based on sidebar start/month count
    fx_months = month_range(start_ym, int(horizon)).strftime("%Y-%m").tolist()
    fx_rates_df["Month"] = fx_rates_df["Month"].astype(
        str) if not fx_rates_df.empty else fx_rates_df.get("Month", [])
    fx_rates_df = fx_rates_df.merge(pd.DataFrame(
        {"Month": fx_months}), on="Month", how="right")
    fx_rates_df["FX_Rate"] = pd.to_numeric(
        fx_rates_df["FX_Rate"], errors="coerce").fillna(1.0)

    with st.form("fx_rates_form"):
        fx_rates_df = st.data_editor(
            fx_rates_df.sort_values("Month"),
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "Month": st.column_config.TextColumn("Month (YYYY-MM)"),
                "FX_Rate": st.column_config.NumberColumn(f"FX Rate (TRY per {fx_currency})", min_value=0.0, step=0.01, format="%.4f"),
            }
        )
        save_fx = st.form_submit_button("Save FX Rates")
    if save_fx:
        st.session_state.fx_rates_df = fx_rates_df
    st.session_state.fx_currency = fx_currency

# -----------------------------
# 2) Base Cost
# -----------------------------
with tab2:
    st.subheader("Base Cost (TRY) + Currency View")

    # Clean any stale computed columns or duplicate names
    base_df_clean = st.session_state.base_cost_df.copy()
    base_df_clean = base_df_clean.loc[:, ~base_df_clean.columns.duplicated()]
    for col in ["Brut_Salary", "Goalpex", "Total_TRY", "Total_Display"]:
        if col in base_df_clean.columns:
            base_df_clean = base_df_clean.drop(columns=[col])
    st.session_state.base_cost_df = base_df_clean

    # FX rates are entered manually here
    currency = st.selectbox("Display Currency", ["TRY", "EUR", "USD"], index=0)
    fx_rate = 1.0
    if currency != "TRY":
        fx_rate = st.number_input(f"{currency} Rate (TRY per {currency})", value=float(
            fx_rate), step=0.01, format="%.4f")

    st.divider()
    if st.button("Add Agent Row"):
        lang_defaults = st.session_state.lang_defaults_df.copy()
        default_lang = lang_defaults["Language"].iloc[0] if not lang_defaults.empty else "DE"
        lang_row = lang_defaults[lang_defaults["Language"] == default_lang]
        new_row = {
            "Language": default_lang,
            "FTE": 1.0,
            "Salary": float(lang_row["Salary"].iloc[0]) if not lang_row.empty else 0.0,
            "OSS": float(lang_row["OSS"].iloc[0]) if not lang_row.empty else st.session_state.config_defaults["OSS"],
            "Food": float(lang_row["Food"].iloc[0]) if not lang_row.empty else st.session_state.config_defaults["Food"],
            "Goalpex_%": float(lang_row["Goalpex_%"].iloc[0]) if not lang_row.empty else st.session_state.config_defaults["Goalpex_%"],
            "Additional_Cost": 0.0,
        }
        st.session_state.base_cost_df = pd.concat(
            [st.session_state.base_cost_df, pd.DataFrame([new_row])],
            ignore_index=True
        )
    st.checkbox("Fill defaults for 0 values", value=True, key="fill_zero_base")

    with st.form("base_cost_form"):
        base_cost_df = st.data_editor(
            st.session_state.base_cost_df,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "Language": st.column_config.SelectboxColumn("Language", options=LANGUAGES),
                "FTE": st.column_config.NumberColumn("FTE", min_value=0.0, step=0.1, format="%.2f"),
                "Salary": st.column_config.NumberColumn("Salary (TRY)", min_value=0.0, step=100.0, format="%.2f"),
                "OSS": st.column_config.NumberColumn("OSS (TRY)", min_value=0.0, step=10.0, format="%.2f"),
                "Food": st.column_config.NumberColumn("Food (TRY)", min_value=0.0, step=10.0, format="%.2f"),
                "Goalpex_%": st.column_config.NumberColumn("Goalpex %", min_value=0.0, step=0.01, format="%.2f"),
                "Additional_Cost": st.column_config.NumberColumn("Additional Cost (TRY)", min_value=0.0, step=10.0, format="%.2f"),
            }
        )
        save_base = st.form_submit_button("Save Base Costs")
    if save_base:
        # Keep only expected columns (avoids duplicates)
        expected_cols = ["Language", "FTE", "Salary",
                         "OSS", "Food", "Goalpex_%", "Additional_Cost"]
        base_cost_df = base_cost_df.loc[:, [
            c for c in expected_cols if c in base_cost_df.columns]]
        base_cost_df = base_cost_df.loc[:, ~base_cost_df.columns.duplicated()]
        # Fill defaults from language config
        lang_defaults = st.session_state.lang_defaults_df.copy()
        lang_defaults["Language"] = lang_defaults["Language"].astype(
            str).str.strip()
        base_cost_df["Language"] = base_cost_df["Language"].astype(
            str).str.strip()
        lang_salary = dict(
            zip(lang_defaults["Language"], lang_defaults["Salary"]))
        lang_oss = dict(zip(lang_defaults["Language"], lang_defaults["OSS"]))
        lang_food = dict(zip(lang_defaults["Language"], lang_defaults["Food"]))
        lang_goalpex = dict(
            zip(lang_defaults["Language"], lang_defaults["Goalpex_%"]))

        fill_zero = st.session_state.get("fill_zero_base", True)
        if fill_zero:
            base_cost_df["Salary"] = base_cost_df["Salary"].replace(0, np.nan)
            base_cost_df["OSS"] = base_cost_df["OSS"].replace(0, np.nan)
            base_cost_df["Food"] = base_cost_df["Food"].replace(0, np.nan)
            base_cost_df["Goalpex_%"] = base_cost_df["Goalpex_%"].replace(
                0, np.nan)

        base_cost_df["Salary"] = base_cost_df["Salary"].fillna(
            base_cost_df["Language"].map(lang_salary)).fillna(0.0)
        base_cost_df["OSS"] = base_cost_df["OSS"].fillna(base_cost_df["Language"].map(
            lang_oss)).fillna(st.session_state.config_defaults["OSS"])
        base_cost_df["Food"] = base_cost_df["Food"].fillna(base_cost_df["Language"].map(
            lang_food)).fillna(st.session_state.config_defaults["Food"])
        base_cost_df["Goalpex_%"] = base_cost_df["Goalpex_%"].fillna(base_cost_df["Language"].map(
            lang_goalpex)).fillna(st.session_state.config_defaults["Goalpex_%"])
        st.session_state.base_cost_df = base_cost_df

    include_benefits = st.checkbox(
        "Include OSS/Food/Additional in Total", value=True, key="include_benefits_base")

    # Calculations
    brut_multiplier = cola_cfg.get("brut_multiplier", 1.58)
    calc_df = base_cost_df.copy()
    calc_df = calc_df.loc[:, ~calc_df.columns.duplicated()]
    for col in ["Brut_Salary", "Goalpex", "Total_TRY", "Total_Display"]:
        if col in calc_df.columns:
            calc_df = calc_df.drop(columns=[col])

    lang_defaults = st.session_state.lang_defaults_df.copy()
    lang_defaults["Language"] = lang_defaults["Language"].astype(
        str).str.strip()
    calc_df["Language"] = calc_df["Language"].astype(str).str.strip()
    lang_salary = dict(zip(lang_defaults["Language"], lang_defaults["Salary"]))
    lang_oss = dict(zip(lang_defaults["Language"], lang_defaults["OSS"]))
    lang_food = dict(zip(lang_defaults["Language"], lang_defaults["Food"]))
    lang_goalpex = dict(
        zip(lang_defaults["Language"], lang_defaults["Goalpex_%"]))

    for col in ["FTE", "Salary", "OSS", "Food", "Goalpex_%", "Additional_Cost"]:
        if col in calc_df.columns:
            calc_df[col] = pd.to_numeric(calc_df[col], errors="coerce")

    calc_df["Salary"] = calc_df["Salary"].fillna(
        calc_df["Language"].map(lang_salary)).fillna(0.0)
    calc_df["OSS"] = calc_df["OSS"].fillna(calc_df["Language"].map(
        lang_oss)).fillna(st.session_state.config_defaults["OSS"])
    calc_df["Food"] = calc_df["Food"].fillna(calc_df["Language"].map(
        lang_food)).fillna(st.session_state.config_defaults["Food"])
    calc_df["Goalpex_%"] = calc_df["Goalpex_%"].fillna(calc_df["Language"].map(
        lang_goalpex)).fillna(st.session_state.config_defaults["Goalpex_%"])
    calc_df["Additional_Cost"] = calc_df["Additional_Cost"].fillna(0.0)
    calc_df["FTE"] = calc_df["FTE"].fillna(0.0)

    calc_df["Goalpex"] = calc_df["Salary"] * calc_df["Goalpex_%"]
    calc_df["Net_Base"] = (
        calc_df["Salary"] +
        calc_df["Goalpex"] +
        calc_df["OSS"] +
        calc_df["Food"] +
        calc_df["Additional_Cost"]
    )
    calc_df["Brut_Base"] = calc_df["Net_Base"] * brut_multiplier

    # Net/Brut totals
    calc_df["Total_Net"] = calc_df["Net_Base"] * calc_df["FTE"]
    calc_df["Total_Brut"] = calc_df["Brut_Base"] * calc_df["FTE"]

    # Keep Total_TRY for downstream usage (use Brut)
    calc_df["Total_TRY"] = calc_df["Total_Brut"]

    if currency == "TRY":
        calc_df["Total_Net_Display"] = calc_df["Total_Net"]
        calc_df["Total_Brut_Display"] = calc_df["Total_Brut"]
    else:
        # TRY per FX -> convert TRY to FX
        calc_df["Total_Net_Display"] = calc_df["Total_Net"] / \
            fx_rate if fx_rate else 0.0
        calc_df["Total_Brut_Display"] = calc_df["Total_Brut"] / \
            fx_rate if fx_rate else 0.0

    # Ensure no duplicate columns before display
    calc_df = calc_df.loc[:, ~calc_df.columns.duplicated()]

    st.subheader("Calculated Cost Summary")
    # Totals per FTE (net)
    calc_df["Salary_Net_Total"] = calc_df["Salary"] * calc_df["FTE"]
    calc_df["OSS_Net_Total"] = calc_df["OSS"] * calc_df["FTE"]
    calc_df["Food_Net_Total"] = calc_df["Food"] * calc_df["FTE"]
    calc_df["Goalpex_Net_Total"] = calc_df["Goalpex"] * calc_df["FTE"]

    display_cols = ["Language", "FTE", "Salary_Net_Total", "OSS_Net_Total",
                    "Food_Net_Total", "Goalpex_Net_Total", "Total_Net_Display", "Total_Brut_Display"]
    display_cols = [c for c in display_cols if c in calc_df.columns]
    display_df = calc_df[display_cols].rename(
        columns={
            "Salary_Net_Total": "Salary Net",
            "OSS_Net_Total": "OSS Net",
            "Food_Net_Total": "Food Net",
            "Goalpex_Net_Total": "Goalpex Net",
            "Total_Net_Display": f"Total Net {currency}",
            "Total_Brut_Display": f"Total Brut {currency}"
        })
    display_df = make_unique_columns(display_df)
    st.dataframe(display_df, use_container_width=True)

    st.session_state["base_cost_calc"] = calc_df

    st.divider()
    st.subheader("Overhead & Operational Cost")
    st.caption("Add operational roles like Operation Manager, Teamleader, Trainer & Quality, RTA, Planner, WFM, Operation Support.")

    role_defaults_df = st.session_state.role_defaults_df.copy()
    role_options = role_defaults_df["Role"].dropna().tolist() if not role_defaults_df.empty else [
        "Operation Manager", "Teamleader", "Trainer & Quality", "RTA", "Planner", "WFM", "Operation Support"
    ]

    if st.button("Add Overhead Row"):
        role_defaults = st.session_state.role_defaults_df.copy()
        default_role = role_defaults["Role"].iloc[0] if not role_defaults.empty else "Operation Manager"
        role_row = role_defaults[role_defaults["Role"] == default_role]
        new_row = {
            "Account": operation_filter if operation_filter != "All" else ACCOUNTS[0],
            "Role": default_role,
            "FTE": 1.0,
            "Salary": float(role_row["Salary"].iloc[0]) if not role_row.empty else 0.0,
            "OSS": float(role_row["OSS"].iloc[0]) if not role_row.empty else st.session_state.config_defaults["OSS"],
            "Food": float(role_row["Food"].iloc[0]) if not role_row.empty else st.session_state.config_defaults["Food"],
            "Goalpex_%": float(role_row["Goalpex_%"].iloc[0]) if not role_row.empty else st.session_state.config_defaults["Goalpex_%"],
            "Additional_Cost": 0.0,
        }
        st.session_state.overhead_cost_df = pd.concat(
            [st.session_state.overhead_cost_df, pd.DataFrame([new_row])],
            ignore_index=True
        )

    st.checkbox("Fill defaults for 0 values",
                value=True, key="fill_zero_overhead")

    with st.form("overhead_cost_form"):
        overhead_df = st.data_editor(
            st.session_state.overhead_cost_df,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "Account": st.column_config.SelectboxColumn("Account", options=ACCOUNTS),
                "Role": st.column_config.SelectboxColumn("Role", options=role_options),
                "FTE": st.column_config.NumberColumn("FTE", min_value=0.0, step=0.1, format="%.2f"),
                "Salary": st.column_config.NumberColumn("Salary (TRY)", min_value=0.0, step=100.0, format="%.2f"),
                "OSS": st.column_config.NumberColumn("OSS (TRY)", min_value=0.0, step=10.0, format="%.2f"),
                "Food": st.column_config.NumberColumn("Food (TRY)", min_value=0.0, step=10.0, format="%.2f"),
                "Goalpex_%": st.column_config.NumberColumn("Goalpex %", min_value=0.0, step=0.01, format="%.2f"),
                "Additional_Cost": st.column_config.NumberColumn("Additional Cost (TRY)", min_value=0.0, step=10.0, format="%.2f"),
            }
        )
        save_overhead = st.form_submit_button("Save Overhead Costs")
    if save_overhead:
        expected_cols = ["Account", "Role", "FTE", "Salary",
                         "OSS", "Food", "Goalpex_%", "Additional_Cost"]
        overhead_df = overhead_df.loc[:, [
            c for c in expected_cols if c in overhead_df.columns]]
        overhead_df = overhead_df.loc[:, ~overhead_df.columns.duplicated()]
        overhead_df["Account"] = overhead_df["Account"].fillna(
            operation_filter if operation_filter != "All" else ACCOUNTS[0]
        )
        overhead_df["Role"] = overhead_df["Role"].fillna("Operation Manager")
        # Fill defaults from role config
        role_defaults = st.session_state.role_defaults_df.copy()
        overhead_df = overhead_df.merge(
            role_defaults, on="Role", how="left", suffixes=("", "_def"))
        fill_zero_overhead = st.session_state.get("fill_zero_overhead", True)
        if fill_zero_overhead:
            for c in ["Salary", "OSS", "Food", "Goalpex_%"]:
                if c in overhead_df.columns:
                    overhead_df[c] = overhead_df[c].replace(0, np.nan)
        overhead_df["Salary"] = overhead_df["Salary"].fillna(
            overhead_df["Salary_def"]).fillna(0.0)
        overhead_df["OSS"] = overhead_df["OSS"].fillna(
            overhead_df["OSS_def"]).fillna(st.session_state.config_defaults["OSS"])
        overhead_df["Food"] = overhead_df["Food"].fillna(
            overhead_df["Food_def"]).fillna(st.session_state.config_defaults["Food"])
        overhead_df["Goalpex_%"] = overhead_df["Goalpex_%"].fillna(
            overhead_df["Goalpex_%_def"]).fillna(st.session_state.config_defaults["Goalpex_%"])
        overhead_df = overhead_df.drop(
            columns=[c for c in overhead_df.columns if c.endswith("_def")])
        st.session_state.overhead_cost_df = overhead_df

    # Overhead calculations
    overhead_calc = st.session_state.overhead_cost_df.copy()
    overhead_calc = overhead_calc.loc[:, ~overhead_calc.columns.duplicated()]
    overhead_calc["Goalpex"] = overhead_calc["Salary"] * \
        overhead_calc["Goalpex_%"]
    overhead_calc["Net_Base"] = (
        overhead_calc["Salary"] +
        overhead_calc["Goalpex"] +
        overhead_calc["OSS"] +
        overhead_calc["Food"] +
        overhead_calc["Additional_Cost"]
    )
    overhead_calc["Brut_Base"] = overhead_calc["Net_Base"] * brut_multiplier

    overhead_calc["Total_Net"] = overhead_calc["Net_Base"] * \
        overhead_calc["FTE"]
    overhead_calc["Total_Brut"] = overhead_calc["Brut_Base"] * \
        overhead_calc["FTE"]

    overhead_calc["Total_TRY"] = overhead_calc["Total_Brut"]

    if currency == "TRY":
        overhead_calc["Total_Net_Display"] = overhead_calc["Total_Net"]
        overhead_calc["Total_Brut_Display"] = overhead_calc["Total_Brut"]
    else:
        overhead_calc["Total_Net_Display"] = overhead_calc["Total_Net"] / \
            fx_rate if fx_rate else 0.0
        overhead_calc["Total_Brut_Display"] = overhead_calc["Total_Brut"] / \
            fx_rate if fx_rate else 0.0

    st.subheader("Overhead Cost Summary")
    overhead_calc["Salary_Net_Total"] = overhead_calc["Salary"] * \
        overhead_calc["FTE"]
    overhead_calc["OSS_Net_Total"] = overhead_calc["OSS"] * \
        overhead_calc["FTE"]
    overhead_calc["Food_Net_Total"] = overhead_calc["Food"] * \
        overhead_calc["FTE"]
    overhead_calc["Goalpex_Net_Total"] = overhead_calc["Goalpex"] * \
        overhead_calc["FTE"]

    overhead_display = overhead_calc[[
        "Account", "Role", "FTE", "Salary_Net_Total", "OSS_Net_Total", "Food_Net_Total",
        "Goalpex_Net_Total", "Total_Net_Display", "Total_Brut_Display"]]
    overhead_display = overhead_display.rename(
        columns={
            "Salary_Net_Total": "Salary Net",
            "OSS_Net_Total": "OSS Net",
            "Food_Net_Total": "Food Net",
            "Goalpex_Net_Total": "Goalpex Net",
            "Total_Net_Display": f"Total Net {currency}",
            "Total_Brut_Display": f"Total Brut {currency}"
        })
    overhead_display = make_unique_columns(overhead_display)
    st.dataframe(overhead_display, use_container_width=True)

    st.session_state["overhead_cost_calc"] = overhead_calc

# -----------------------------
# 3) Monthly Plan
# -----------------------------
with tab3:
    st.subheader("Monthly Plan (Volume + Agent Cost)")
    st.caption(
        "Volume = hours — you decide. Agent Cost = costs for that month row.")

    st.subheader("FTE Import")
    st.caption(
        "Download the template, fill FTE per month/account/language, then upload to update the plan.")

    # Template uses current months and accounts
    # Build template from months + selected operation + languages from Base Cost
    months_list = sorted(
        st.session_state.plan_df["Month"].astype(str).unique().tolist())
    op_list = [operation_filter] if operation_filter != "All" else ACCOUNTS
    base_langs = []
    if "base_cost_calc" in st.session_state:
        base_langs = sorted(
            st.session_state["base_cost_calc"]["Language"].astype(str).unique().tolist())
    if not base_langs:
        base_langs = sorted(
            st.session_state.plan_df["Language"].astype(str).unique().tolist())

    template_rows = []
    for op in op_list:
        for m in months_list:
            for lang in base_langs:
                template_rows.append(
                    {"Month": m, "Account": op, "Language": lang})
    template_df = pd.DataFrame(template_rows)
    template_df["FTE"] = 0.0
    template_df["Production_Hours"] = 0.0
    template_bytes = df_to_excel_bytes(template_df, sheet_name="FTE_Template")
    st.download_button(
        "⬇️ Download FTE Template (xlsx)",
        data=template_bytes,
        file_name="fte_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    fte_file = st.file_uploader("Upload FTE/Hours File (xlsx)", type=["xlsx"])
    if fte_file is not None:
        try:
            fte_df = pd.read_excel(fte_file)
            # Expect columns: Month, Account, Language, FTE, Production_Hours
            required_cols = {"Month", "Account", "Language"}
            if not required_cols.issubset(set(fte_df.columns)):
                st.error(
                    "File must include columns: Month, Account, Language (plus optional FTE, Production_Hours)")
            else:
                plan_df = st.session_state.plan_df.copy()
                fte_df["Month"] = fte_df["Month"].astype(str)
                fte_df["Account"] = fte_df["Account"].astype(str)
                fte_df["Language"] = fte_df["Language"].astype(str)
                if "FTE" in fte_df.columns:
                    fte_df["FTE"] = pd.to_numeric(
                        fte_df["FTE"], errors="coerce").fillna(0.0)
                if "Production_Hours" in fte_df.columns:
                    fte_df["Production_Hours"] = pd.to_numeric(
                        fte_df["Production_Hours"], errors="coerce").fillna(0.0)

                cols = ["Month", "Account", "Language"]
                if "FTE" in fte_df.columns:
                    cols.append("FTE")
                if "Production_Hours" in fte_df.columns:
                    cols.append("Production_Hours")

                plan_df = plan_df.merge(
                    fte_df[cols],
                    on=["Month", "Account", "Language"],
                    how="left",
                    suffixes=("", "_new")
                )
                if "FTE_new" in plan_df.columns:
                    plan_df["FTE"] = plan_df["FTE_new"].combine_first(
                        plan_df["FTE"])
                    plan_df = plan_df.drop(columns=["FTE_new"])
                if "Production_Hours_new" in plan_df.columns:
                    plan_df["Production_Hours"] = plan_df["Production_Hours_new"].combine_first(
                        plan_df["Production_Hours"])
                    plan_df = plan_df.drop(columns=["Production_Hours_new"])
                st.session_state.plan_df = plan_df
                st.success("FTE/Production Hours imported into Monthly Plan.")
        except Exception as e:
            st.error(f"Could not read FTE file: {e}")

    if "base_cost_calc" in st.session_state:
        if st.button("Apply Base Cost by Language (monthly)"):
            if operation_filter == "All":
                st.warning(
                    "Select a specific Operation in the sidebar to apply Agent Cost, to avoid copying the same cost to all operations.")
                # Do not apply when All is selected
                st.stop()
            plan_df = st.session_state.plan_df.copy()
            base_calc = st.session_state["base_cost_calc"].copy()

            # Ensure plan has rows for all languages in base cost for this operation
            op_months = sorted(plan_df.loc[plan_df["Account"] == operation_filter, "Month"].astype(
                str).unique().tolist())
            base_langs = sorted(
                base_calc["Language"].astype(str).unique().tolist())
            existing_keys = set(
                zip(
                    plan_df["Month"].astype(str),
                    plan_df["Account"].astype(str),
                    plan_df["Language"].astype(str),
                )
            )
            new_rows = []
            for m in op_months:
                for lang in base_langs:
                    key = (m, operation_filter, lang)
                    if key not in existing_keys:
                        new_rows.append({
                            "Month": m,
                            "Account": operation_filter,
                            "Language": lang,
                            "Production_Hours": 0.0,
                            "FTE": 0.0,
                            "Base_Cost": 0.0,
                            "Overhead_Cost": 0.0,
                            "Notes": ""
                        })
            if new_rows:
                plan_df = pd.concat(
                    [plan_df, pd.DataFrame(new_rows)], ignore_index=True)

            per_lang = base_calc.groupby("Language", as_index=False)[
                "Total_Brut"].sum()
            plan_df = plan_df.merge(per_lang, on="Language", how="left")
            plan_df["Total_Brut"] = plan_df["Total_Brut"].fillna(0.0)

            # If multiple rows per Month+Language, split cost equally
            counts = plan_df.groupby(["Month", "Language", "Account"])[
                "Language"].transform("count")
            counts = counts.replace(0, 1)
            new_base = plan_df["Total_Brut"] / counts
            if operation_filter != "All":
                mask = plan_df["Account"] == operation_filter
                # Scale by monthly FTE vs base FTE from base_cost_calc
                base_fte_map = base_calc.set_index("Language")["FTE"].to_dict()
                base_fte = plan_df["Language"].map(base_fte_map).fillna(1.0)
                scale = plan_df["FTE"] / base_fte.replace(0, 1.0)
                plan_df.loc[mask, "Base_Cost"] = new_base[mask] * scale[mask]
            else:
                base_fte_map = base_calc.set_index("Language")["FTE"].to_dict()
                base_fte = plan_df["Language"].map(base_fte_map).fillna(1.0)
                scale = plan_df["FTE"] / base_fte.replace(0, 1.0)
                plan_df["Base_Cost"] = new_base * scale
            plan_df = plan_df.drop(columns=["Total_Brut"])
            st.session_state.plan_df = plan_df
            st.success("Base Cost applied to plan by Language.")

    if "overhead_cost_calc" in st.session_state:
        if st.button("Apply Overhead by Operation (monthly)"):
            plan_df = st.session_state.plan_df.copy()
            overhead_calc = st.session_state["overhead_cost_calc"].copy()
            per_op = overhead_calc.groupby("Account", as_index=False)[
                "Total_Brut"].sum()
            plan_df = plan_df.merge(per_op, on="Account", how="left")
            plan_df["Total_Brut"] = plan_df["Total_Brut"].fillna(0.0)

            # Distribute overhead across rows per Month+Account
            counts = plan_df.groupby(["Month", "Account"])[
                "Account"].transform("count")
            counts = counts.replace(0, 1)
            if "Overhead_Cost" not in plan_df.columns:
                plan_df["Overhead_Cost"] = 0.0
            plan_df["Overhead_Cost"] = plan_df["Total_Brut"] / counts
            plan_df = plan_df.drop(columns=["Total_Brut"])
            st.session_state.plan_df = plan_df
            st.success("Overhead applied to plan by Operation.")

    plan_view = st.session_state.plan_df
    if operation_filter != "All":
        plan_view = plan_view[plan_view["Account"] == operation_filter]

    with st.form("plan_form"):
        plan_df = st.data_editor(
            plan_view,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "Month": st.column_config.TextColumn("Month (YYYY-MM)"),
                "Account": st.column_config.SelectboxColumn("Account", options=ACCOUNTS),
                "Language": st.column_config.SelectboxColumn("Language", options=LANGUAGES),
                "Production_Hours": st.column_config.NumberColumn("Production Hours", min_value=0.0, step=1.0, format="%.2f"),
                "FTE": st.column_config.NumberColumn("FTE", min_value=0.0, step=0.1, format="%.2f"),
                "Base_Cost": st.column_config.NumberColumn("Agent Cost", min_value=0.0, step=10.0, format="%.2f"),
                "Overhead_Cost": st.column_config.NumberColumn("Overhead Cost", min_value=0.0, step=10.0, format="%.2f"),
                "Notes": st.column_config.TextColumn("Notes"),
            }
        )
        save_plan = st.form_submit_button("Save Monthly Plan")
    if save_plan:
        if operation_filter != "All":
            full = st.session_state.plan_df.copy()
            full = full[full["Account"] != operation_filter]
            st.session_state.plan_df = pd.concat(
                [full, plan_df], ignore_index=True)
        else:
            st.session_state.plan_df = plan_df

# -----------------------------
# 4) Results
# -----------------------------
with tab4:
    st.subheader("Results")

    prices_df = st.session_state.prices_df.copy()
    plan_df = st.session_state.plan_df.copy()

    # ensure numeric
    for col in ["UnitPrice"]:
        if col in prices_df.columns:
            prices_df[col] = pd.to_numeric(
                prices_df[col], errors="coerce").fillna(0.0)

    for col in ["Production_Hours", "FTE", "Base_Cost", "Overhead_Cost"]:
        if col in plan_df.columns:
            plan_df[col] = pd.to_numeric(
                plan_df[col], errors="coerce").fillna(0.0)

    result_df = compute_budget(prices_df, plan_df, cola_cfg,
                               st.session_state.fx_rates_df if "fx_rates_df" in st.session_state else None)

    if operation_filter != "All":
        result_df = result_df[result_df["Account"] == operation_filter]

    # KPIs
    total_revenue = result_df["Revenue"].sum()
    total_base_cost = result_df["Base_Cost"].sum(
    ) if "Base_Cost" in result_df.columns else 0.0
    total_overhead_cost = result_df["Overhead_Cost"].sum(
    ) if "Overhead_Cost" in result_df.columns else 0.0
    total_cost = result_df["Adj_Cost"].sum()
    total_gm = result_df["GM"].sum()
    total_gm_pct = 0 if total_revenue == 0 else total_gm / total_revenue

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Revenue", f"{total_revenue:,.2f}")
    c2.metric("Agent Cost (Base)", f"{total_base_cost:,.2f}")
    c3.metric("Overhead Cost", f"{total_overhead_cost:,.2f}")
    c4.metric("Total Cost", f"{total_cost:,.2f}")
    c5, c6 = st.columns(2)
    c5.metric("Gross Margin", f"{total_gm:,.2f}")
    c6.metric("GM %", f"{total_gm_pct*100:,.2f}%")

    st.divider()

    st.write("Detailed Results")
    result_display = result_df.copy()
    if "Notes" in result_display.columns:
        result_display = result_display.drop(columns=["Notes"])
    if "Month_dt" in result_display.columns:
        result_display = result_display.drop(columns=["Month_dt"])

    # Arrange columns for visibility
    ordered_cols = [
        "Month", "Account", "Language", "Production_Hours", "FTE",
        "Base_Cost", "Overhead_Cost", "Adj_Cost",
        "Unit_Price", "COLA_%", "Adj_Unit_Price", "Billing_Mode",
        "FX_Rate", "Eff_Production_Hours", "Eff_FTE", "Eff_FTE_Hours",
        "Billable_Hours", "Revenue", "GM", "GM_%"
    ]
    ordered_cols = [c for c in ordered_cols if c in result_display.columns]
    result_display = result_display[ordered_cols]
    # Add percent display columns
    if "COLA_%" in result_display.columns:
        result_display["COLA_%_Display"] = (
            result_display["COLA_%"] * 100).round(2).astype(str) + "%"
        result_display = result_display.drop(columns=["COLA_%"])
    if "GM_%" in result_display.columns:
        result_display["GM_%_Display"] = (
            result_display["GM_%"] * 100).round(2).astype(str) + "%"
        result_display = result_display.drop(columns=["GM_%"])

    result_display = result_display.rename(columns={
        "Base_Cost": "Agent Cost",
        "Overhead_Cost": "Overhead Cost",
        "Adj_Cost": "Total Cost",
        "Adj_Unit_Price": "Adj. Unit Price",
        "COLA_%_Display": "COLA_%",
        "GM_%_Display": "GM_%",
    })

    # Reorder after display/rename so COLA_% is between Unit_Price and Adj. Unit Price
    ordered_display_cols = [
        "Month", "Account", "Language", "Production_Hours", "FTE",
        "Agent Cost", "Overhead Cost", "Total Cost",
        "Unit_Price", "COLA_%", "Adj. Unit Price", "Billing_Mode",
        "FX_Rate", "Eff_Production_Hours", "Eff_FTE", "Eff_FTE_Hours",
        "Billable_Hours", "Revenue", "GM", "GM_%"
    ]
    ordered_display_cols = [
        c for c in ordered_display_cols if c in result_display.columns]
    result_display = result_display[ordered_display_cols]

    # Keep a numeric copy for export
    result_export = result_display.copy()

    # Format amounts with thousand separators for display
    money_cols = [
        "Agent Cost", "Overhead Cost", "Total Cost", "Unit_Price",
        "Adj. Unit Price", "FX_Rate", "Eff_Production_Hours",
        "Eff_FTE", "Eff_FTE_Hours", "Billable_Hours", "Revenue", "GM"
    ]
    result_display = format_number_cols(result_display, money_cols, decimals=2)
    if "FX_Rate" in result_display.columns:
        result_display = format_number_cols(result_display, ["FX_Rate"], decimals=4)

    st.dataframe(
        result_display,
        use_container_width=True,
        column_config={
            "Production_Hours": st.column_config.TextColumn("Production_Hours"),
            "FTE": st.column_config.TextColumn("FTE"),
            "Agent Cost": st.column_config.TextColumn("Agent Cost"),
            "Overhead Cost": st.column_config.TextColumn("Overhead_Cost"),
            "Total Cost": st.column_config.TextColumn("Total Cost"),
            "Unit_Price": st.column_config.TextColumn("Unit_Price"),
            "COLA_%": st.column_config.TextColumn("COLA_%"),
            "Adj. Unit Price": st.column_config.TextColumn("Adj. Unit Price"),
            "FX_Rate": st.column_config.TextColumn("FX_Rate"),
            "Eff_Production_Hours": st.column_config.TextColumn("Eff_Production_Hours"),
            "Eff_FTE": st.column_config.TextColumn("Eff_FTE"),
            "Eff_FTE_Hours": st.column_config.TextColumn("Eff_FTE_Hours"),
            "Billable_Hours": st.column_config.TextColumn("Billable_Hours"),
            "Revenue": st.column_config.TextColumn("Revenue"),
            "GM": st.column_config.TextColumn("GM"),
            "GM_%": st.column_config.TextColumn("GM_%"),
        }
    )

    # Pivot summary
    st.write("Summary by Account")
    if result_df.empty:
        st.info("No rows for the selected operation.")
        pivot_acc = pd.DataFrame(
            columns=["Account", "Revenue", "Adj_Cost", "GM", "GM_%"])
        st.dataframe(pivot_acc, use_container_width=True)
    else:
        pivot_acc = result_df.pivot_table(
            index="Account",
            values=["Revenue", "Adj_Cost", "GM"],
            aggfunc="sum"
        ).reset_index()
        if "Revenue" in pivot_acc.columns and "GM" in pivot_acc.columns:
            pivot_acc["GM_%"] = np.where(
                pivot_acc["Revenue"] == 0, 0, pivot_acc["GM"] / pivot_acc["Revenue"])
        else:
            pivot_acc["GM_%"] = 0.0
        pivot_acc_display = pivot_acc.copy()
        pivot_acc_display["GM_%"] = (pivot_acc_display["GM_%"] * 100).round(2).astype(str) + "%"
        pivot_acc_display = format_number_cols(pivot_acc_display, ["Revenue", "Adj_Cost", "GM"], decimals=2)
        st.dataframe(
            pivot_acc_display,
            use_container_width=True,
            column_config={
                "Revenue": st.column_config.TextColumn("Revenue"),
                "Adj_Cost": st.column_config.TextColumn("Adj Cost"),
                "GM": st.column_config.TextColumn("GM"),
                "GM_%": st.column_config.TextColumn("GM %"),
            }
        )

    st.write("Summary by Account & Language")
    if result_df.empty:
        pivot_acc_lang = pd.DataFrame(
            columns=["Account", "Language", "Revenue", "Adj_Cost", "GM", "GM_%"])
        st.dataframe(pivot_acc_lang, use_container_width=True)
    else:
        pivot_acc_lang = result_df.pivot_table(
            index=["Account", "Language"],
            values=["Revenue", "Adj_Cost", "GM"],
            aggfunc="sum"
        ).reset_index()
        if "Revenue" in pivot_acc_lang.columns and "GM" in pivot_acc_lang.columns:
            pivot_acc_lang["GM_%"] = np.where(
                pivot_acc_lang["Revenue"] == 0, 0, pivot_acc_lang["GM"] / pivot_acc_lang["Revenue"])
        else:
            pivot_acc_lang["GM_%"] = 0.0
        pivot_acc_lang_display = pivot_acc_lang.copy()
        pivot_acc_lang_display["GM_%"] = (pivot_acc_lang_display["GM_%"] * 100).round(2).astype(str) + "%"
        pivot_acc_lang_display = format_number_cols(pivot_acc_lang_display, ["Revenue", "Adj_Cost", "GM"], decimals=2)
        st.dataframe(
            pivot_acc_lang_display,
            use_container_width=True,
            column_config={
                "Revenue": st.column_config.TextColumn("Revenue"),
                "Adj_Cost": st.column_config.TextColumn("Adj Cost"),
                "GM": st.column_config.TextColumn("GM"),
                "GM_%": st.column_config.TextColumn("GM %"),
            }
        )

    # Downloads (XLSX)
    xlsx = df_to_excel_bytes(result_export, sheet_name="Detailed Results")
    st.download_button("⬇️ Download detailed XLSX", data=xlsx,
                       file_name="budget_results_detailed.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    xlsx2 = df_to_excel_bytes(pivot_acc, sheet_name="Summary by Account")
    st.download_button("⬇️ Download account summary XLSX", data=xlsx2,
                       file_name="budget_results_summary_account.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    xlsx3 = df_to_excel_bytes(
        pivot_acc_lang, sheet_name="Summary by Account & Language")
    st.download_button("⬇️ Download account+language summary XLSX", data=xlsx3,
                       file_name="budget_results_summary_account_language.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
