import io
import re
import threading
import time
from datetime import datetime

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials
from gspread.utils import rowcol_to_a1

st.set_page_config(page_title="Bank Product Database", page_icon="🏦", layout="wide")

CORE = ["Bank Category", "Bank", "Account Type", "Product Name"]
META = ["Added By", "Added At"]
FIXED = CORE + META
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
NEW = "➕ Add new..."


# ---------------------------------------------------------------- access gate
def gate():
    """Optional shared passcode (set app_password in secrets). Streamlit Cloud's
    private-app invite list is the main protection; this is a second layer."""
    pw = st.secrets.get("app_password", "")
    if not pw or st.session_state.get("ok"):
        return
    st.title("🏦 Bank Product Database")
    entered = st.text_input("Team passcode", type="password")
    if entered:
        if entered == pw:
            st.session_state["ok"] = True
            st.rerun()
        else:
            st.error("Wrong passcode")
    st.stop()


# ---------------------------------------------------------------- sheet access
@st.cache_resource
def get_worksheet():
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]), scopes=SCOPES
    )
    client = gspread.authorize(creds)
    sh = client.open_by_key(st.secrets["sheet_id"])
    try:
        ws = sh.worksheet("Products")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet("Products", rows=1000, cols=60)
    if not ws.row_values(1):
        ws.update(range_name="A1", values=[FIXED])
    return ws


@st.cache_data(ttl=10, show_spinner=False)
def load_df() -> pd.DataFrame:
    ws = get_worksheet()
    values = ws.get_all_values()
    if not values:
        return pd.DataFrame(columns=FIXED)
    header, rows = values[0], values[1:]
    df = pd.DataFrame(rows, columns=header)
    for c in FIXED:
        if c not in df.columns:
            df[c] = ""
    return df


@st.cache_resource
def save_lock():
    """One lock shared by every user session of this app (Streamlit runs them in one process)."""
    return threading.Lock()


def _row_number(res):
    """Sheet row the product landed in, from Google's append response (None if unavailable)."""
    try:
        rng = res["updates"]["updatedRange"]  # e.g. "Products!A42:K42"
        return int(re.search(r"[A-Z]+(\d+)", rng.split("!")[-1]).group(1))
    except Exception:
        return None


def save_row(row: dict):
    """Append one product and return the sheet row number it landed in.
    New features become new columns at the far right. Saves are serialised with a lock and the
    header is re-read inside it, so simultaneous saves from different windows cannot clash.
    Temporary Google errors (rate limit / server hiccup) are retried automatically."""
    ws = get_worksheet()
    last_err = None
    for attempt in range(3):
        try:
            with save_lock():
                header = ws.row_values(1)  # read INSIDE the lock: always the latest columns
                lower = {h.lower(): h for h in header}
                row = {lower.get(k.lower(), k): v for k, v in row.items()}  # reuse existing spelling
                new_cols = [c for c in row if c not in header]
                if new_cols:
                    start = len(header) + 1
                    if start + len(new_cols) - 1 > ws.col_count:
                        ws.add_cols(len(new_cols) + 10)
                    # write only the NEW header cells, never the whole header row
                    ws.update(range_name=rowcol_to_a1(1, start), values=[new_cols])
                    header = header + new_cols
                res = ws.append_row(
                    [row.get(h, "") for h in header],
                    value_input_option="USER_ENTERED",
                    table_range="A1",
                )
            load_df.clear()
            return _row_number(res)
        except gspread.exceptions.APIError as e:
            code = getattr(getattr(e, "response", None), "status_code", 0)
            if code not in (429, 500, 502, 503, 504):
                raise
            last_err = e
            time.sleep(2 * (attempt + 1))  # wait outside the lock so others can save meanwhile
    raise last_err


def to_excel(df: pd.DataFrame) -> bytes:
    feats = [c for c in df.columns if c not in FIXED]
    out = df[CORE + feats + META]
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        out.to_excel(xw, index=False, sheet_name="Products")
        ws = xw.sheets["Products"]
        ws.freeze_panes = "E2"
        for i, col in enumerate(out.columns, start=1):
            ws.column_dimensions[ws.cell(1, i).column_letter].width = max(14, min(32, len(col) + 4))
    return buf.getvalue()


# ---------------------------------------------------------------- helpers
def recent_first(series: pd.Series) -> list:
    """Unique non-empty values, most recently added first."""
    vals = [v for v in series.tolist() if str(v).strip()]
    return list(dict.fromkeys(reversed(vals)))


def feature_cols(df):
    return [c for c in df.columns if c not in FIXED]


def pick(label, options, key, default=None):
    """Selectbox of recent values + 'Add new' option that reveals a text box."""
    opts = options + [NEW]
    idx = opts.index(default) if default in opts else 0
    v = st.session_state.get("form_ver", 0)
    choice = st.selectbox(label, opts, index=idx, key=f"{key}_sel_{v}")
    if choice == NEW:
        return st.text_input(f"New {label.lower()}", key=f"{key}_new_{v}").strip()
    return choice


# ---------------------------------------------------------------- app
gate()
st.title("🏦 Bank Product Database")

try:
    df = load_df()
except Exception as e:
    st.error(
        "Could not connect to Google Sheets. Check `sheet_id`, the service-account secrets, "
        "and that the sheet is shared with the service-account email as Editor."
    )
    st.exception(e)
    st.stop()

feats = feature_cols(df)


# ---------------------------------------------------------------- who is working?
def known_members(frame):
    names = list(st.secrets.get("members", [])) + recent_first(frame["Added By"])
    seen, out = set(), []
    for n in names:
        n = str(n).strip()
        if n and n.lower() not in seen:
            seen.add(n.lower())
            out.append(n)
    return sorted(out, key=str.lower)


def feature_origins(frame, feature_list):
    """Who first filled each feature (so teammates can see who introduced it)."""
    out = {}
    for f in feature_list:
        hit = frame.index[frame[f].astype(str).str.strip() != ""]
        out[f] = frame.loc[hit[0], "Added By"] if len(hit) else ""
    return out


with st.sidebar:
    st.header("👤 Your workspace")
    members = known_members(df)
    qp_name = str(st.query_params.get("member", "")).strip()
    if qp_name and qp_name.lower() not in [m.lower() for m in members]:
        members = sorted(members + [qp_name], key=str.lower)
    start = next((i for i, m in enumerate(members) if m.lower() == qp_name.lower()), None) if qp_name else None
    me = st.selectbox(
        "Who are you?", members, index=start, accept_new_options=True,
        placeholder="Pick or type your name", key="member_sel",
    )
    me = (me or "").strip()
    if me and me != qp_name:
        st.query_params["member"] = me  # makes a personal link: ...app-url/?member=YourName
    if me:
        st.success(f"Working as **{me}**")
        st.caption(
            "Bookmark this page: the link remembers your name. Everyone works in their own window "
            "but shares one feature list and one Google Sheet."
        )
    else:
        st.info("Pick your name to start (type it if you're new).")

mine = df[df["Added By"].str.lower() == me.lower()] if me else df.iloc[0:0]
origins = feature_origins(df, feats)
latest = feats[-5:][::-1]  # right-most columns are the newest features
new_from_others = [f for f in latest if me and origins.get(f, "").strip().lower() not in ("", me.lower())]

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Team products", len(df))
c2.metric("My products", len(mine) if me else "-")
c3.metric("Banks", df["Bank"].replace("", pd.NA).nunique())
c4.metric("Features tracked", len(feats))
c5.metric("Account types", df["Account Type"].replace("", pd.NA).nunique())

tab_add, tab_mine, tab_browse, tab_overview = st.tabs(
    ["➕ Add product", "👤 My work", "🔎 All data & download", "📊 Team overview"]
)

# ------------------------------------------------------------------ ADD
with tab_add:
    ver = st.session_state.setdefault("form_ver", 0)
    if "flash" in st.session_state:
        st.success(st.session_state.pop("flash"))
    left, right = st.columns([1, 1.4], gap="large")

    with left:
        st.subheader("Product details")
        if me:
            st.caption(f"Adding as **{me}**. Your banks are listed first.")
        else:
            st.warning("Pick your name in the sidebar before saving.")

        my_banks = recent_first(mine["Bank"])
        banks = my_banks + [b for b in recent_first(df["Bank"]) if b not in my_banks]
        bank = pick("Bank", banks, "bank")

        # category auto-fills from the bank's earlier entries
        known_cat = None
        if bank and bank in set(df["Bank"]):
            known_cat = df.loc[df["Bank"] == bank, "Bank Category"].iloc[-1] or None
        categories = recent_first(df["Bank Category"])
        category = pick("Bank Category", categories, f"cat_{bank}", default=known_cat)

        acc_types = recent_first(df["Account Type"])
        acc_type = pick("Account Type", acc_types, "acc")

        product = st.text_input("Product Name", key=f"product_{ver}").strip()

        if bank and product:
            dup = df[(df["Bank"] == bank) & (df["Product Name"].str.lower() == product.lower())]
            if len(dup):
                st.warning("This bank already has a product with this name.")

        if bank and bank in set(df["Bank"]):
            st.caption(f"{bank} already has {int((df['Bank'] == bank).sum())} product(s) in the database.")

    with right:
        st.subheader("Features")
        st.caption(
            f"{len(feats)} features already listed (A-Z). Pick a recent value from the dropdown or type a new one; "
            "leave blank to skip."
        )
        if latest:
            with st.expander("🆕 Latest features added by the team"):
                for f in latest:
                    st.write(f"• **{f}**, first used by {origins.get(f) or 'unknown'}")
        feat_values = {}
        if feats:
            with st.container(height=520, border=True):
                cols = st.columns(2)
                for i, f in enumerate(sorted(feats, key=str.lower)):  # A-Z, new features slot in automatically
                    with cols[i % 2]:
                        v = st.selectbox(
                            ("🆕 " + f) if f in new_from_others else f,
                            recent_first(df[f]),  # most recently used values first
                            index=None,
                            accept_new_options=True,
                            placeholder="select or type...",
                            key=f"fv_{f}_{ver}",
                        )
                    if v and str(v).strip():
                        feat_values[f] = str(v).strip()
        else:
            st.info("No features yet. Add the first ones below.")

        st.markdown("**➕ New features** (not in the list above): add a row per feature")
        new_edit = st.data_editor(
            pd.DataFrame({"Feature": [], "Value": []}, dtype="object"),
            num_rows="dynamic",
            width="stretch",
            hide_index=True,
            column_config={
                "Feature": st.column_config.TextColumn("Feature", help="e.g. insurance_25000"),
                "Value": st.column_config.TextColumn("Value", help="e.g. Yes / No / 5% / 25000"),
            },
            key=f"newfeat_{ver}",
        )

    st.divider()
    if st.button("💾 Save product", type="primary", width="stretch", disabled=not me):
        problems = []
        if not (bank and category and acc_type and product):
            problems.append("Bank, Bank Category, Account Type and Product Name are all required.")
        if problems:
            for p in problems:
                st.error(p)
        else:
            row = {
                "Bank Category": category,
                "Bank": bank,
                "Account Type": acc_type,
                "Product Name": product,
                "Added By": me,
                "Added At": datetime.now().strftime("%Y-%m-%d %H:%M"),
            }
            row.update(feat_values)
            new_feats = 0
            lower = {x.lower(): x for x in feats}
            for _, r in new_edit.iterrows():
                f = str(r["Feature"] or "").strip()
                v = str(r["Value"] or "").strip()
                if f and v:
                    f = lower.get(f.lower(), f)  # reuse existing spelling/case if it already exists
                    if f in FIXED:
                        st.error(f"'{f}' is a reserved column name; rename that feature.")
                        st.stop()
                    new_feats += f not in feats
                    row[f] = v
            try:
                with st.spinner("Saving to the shared sheet..."):
                    row_no = save_row(row)
            except Exception as e:
                st.error(
                    "Could not save right now, so nothing was added. Your entries are still on the screen: "
                    f"wait a few seconds and click Save again. ({type(e).__name__})"
                )
                st.stop()
            st.session_state["flash"] = (
                f"Saved **{product}** ({bank}) with {len(row) - len(FIXED)} feature values"
                + (f", including {new_feats} new feature(s)" if new_feats else "")
                + (f". It is in row {row_no} of the Google Sheet." if row_no else ".")
            )
            st.session_state["form_ver"] += 1  # fresh form; bank list now shows this bank first
            st.rerun()

# ------------------------------------------------------------------ MY WORK
with tab_mine:
    if not me:
        st.info("Pick your name in the sidebar to see your own products.")
    elif mine.empty:
        st.info("You haven't added anything yet. Use the **Add product** tab.")
    else:
        st.subheader(f"{me}'s products")
        filled = (
            mine[feats].astype(str).apply(lambda c: c.str.strip().ne("")).sum(axis=1)
            if feats else pd.Series(0, index=mine.index)
        )
        m1, m2, m3 = st.columns(3)
        m1.metric("My products", len(mine))
        m2.metric("My banks", mine["Bank"].nunique())
        m3.metric("Avg features filled", f"{filled.mean():.1f} / {len(feats)}")
        mv = mine[CORE + feats + META].copy()
        mv.insert(4, "Features filled", filled.values)
        st.dataframe(mv, width="stretch", height=420)
        st.download_button(
            "⬇️ Download my products (Excel)",
            data=to_excel(mine),
            file_name=f"{me}_products_{datetime.now():%Y%m%d_%H%M}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.caption("Low 'Features filled' on a product usually means values are still missing. "
                   "Fix them directly in the Google Sheet.")

# ------------------------------------------------------------------ BROWSE
with tab_browse:
    st.subheader("All products")
    f1, f2, f3, f4 = st.columns(4)
    sel_cat = f1.multiselect("Bank Category", sorted(df["Bank Category"].unique()))
    sel_bank = f2.multiselect("Bank", sorted(df["Bank"].unique()))
    sel_acc = f3.multiselect("Account Type", sorted(df["Account Type"].unique()))
    sel_by = f4.multiselect("Added By", sorted(df["Added By"].unique()))
    view = df
    if sel_by:
        view = view[view["Added By"].isin(sel_by)]
    if sel_cat:
        view = view[view["Bank Category"].isin(sel_cat)]
    if sel_bank:
        view = view[view["Bank"].isin(sel_bank)]
    if sel_acc:
        view = view[view["Account Type"].isin(sel_acc)]
    st.dataframe(view[CORE + feats + META], width="stretch", height=460)

    d1, d2, _ = st.columns([1, 1, 2])
    d1.download_button(
        "⬇️ Download Excel (all data)",
        data=to_excel(df),
        file_name=f"bank_products_{datetime.now():%Y%m%d_%H%M}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
    if d2.button("🔄 Refresh"):
        load_df.clear()
        st.rerun()
    st.caption("To fix or delete a row, edit it directly in the Google Sheet — the app picks changes up within seconds.")

# ------------------------------------------------------------------ OVERVIEW
with tab_overview:
    o1, o2 = st.columns(2)
    with o1:
        st.subheader("Products per bank")
        if len(df):
            st.bar_chart(df[df["Bank"] != ""].groupby("Bank").size().sort_values(ascending=False))
    with o2:
        st.subheader("Products per category")
        if len(df):
            st.bar_chart(df[df["Bank Category"] != ""].groupby("Bank Category").size())

    st.subheader(f"Feature coverage ({len(feats)} features)")
    if feats:
        filled = (df[feats].replace("", pd.NA).notna()).sum()
        cov = pd.DataFrame(
            {"Feature": filled.index, "Products with value": filled.values,
             "Coverage %": (100 * filled.values / max(len(df), 1)).round(1)}
        ).sort_values("Products with value", ascending=False)
        st.dataframe(cov, width="stretch", hide_index=True, height=420)
    else:
        st.info("No features yet — add your first product.")

    st.subheader("Team progress")
    if len(df):
        st.bar_chart(df[df["Added By"] != ""].groupby("Added By").size())

    st.subheader("Contributions")
    if len(df):
        st.dataframe(
            df.groupby(["Added By", "Bank"]).size().rename("Products").reset_index(),
            width="stretch", hide_index=True,
        )
