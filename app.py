"""
Material Delivery Schedule Alert Dashboard
Run: streamlit run app.py
"""
import io
import tempfile
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

ROOT = Path(__file__).parent
DEFAULT_DATA_DIR = ROOT / "data"

import sys
sys.path.insert(0, str(ROOT / "src"))
from parser import load_all_files, parse_esp, parse_suborder_list, _is_esp_file

# ── page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="材料到料時程警示儀表板",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
  .alert-red   { background:#ffe0e0; border-left:4px solid #d32f2f; padding:8px 12px; border-radius:4px; margin:4px 0; }
  .alert-orange{ background:#fff3e0; border-left:4px solid #f57c00; padding:8px 12px; border-radius:4px; margin:4px 0; }
  .alert-green { background:#e8f5e9; border-left:4px solid #388e3c; padding:8px 12px; border-radius:4px; margin:4px 0; }
  .week-header { font-size:1.1rem; font-weight:700; color:#1565c0; margin-top:1rem; }
  .upload-box  { background:#e3f2fd; border-radius:8px; padding:12px; margin-bottom:8px; }
</style>
""", unsafe_allow_html=True)

# ── sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.title("📦 材料管控儀表板")
    st.divider()

    # ── 資料來源 ──────────────────────────────────────────────────────────────
    st.markdown("### 📂 資料來源")
    data_mode = st.radio(
        "選擇資料來源",
        ["⬆️ 上傳 Excel 檔案", "📁 本地資料夾"],
        index=0,
        label_visibility="collapsed",
    )

    uploaded_files = []
    data_dir = str(DEFAULT_DATA_DIR)

    if data_mode == "⬆️ 上傳 Excel 檔案":
        st.markdown('<div class="upload-box">', unsafe_allow_html=True)
        uploaded_files = st.file_uploader(
            "拖拉或點選上傳 Excel 報告",
            type=["xlsx", "xls"],
            accept_multiple_files=True,
            help="支援 ESP 報告與廠商 Sub-Order List，可一次上傳多個檔案",
        )
        st.markdown("</div>", unsafe_allow_html=True)
        if uploaded_files:
            st.success(f"已上傳 {len(uploaded_files)} 個檔案")
            for f in uploaded_files:
                st.caption(f"• {f.name}")
    else:
        data_dir = st.text_input("資料夾路徑", str(DEFAULT_DATA_DIR))
        st.caption("支援 ESP 報告與廠商 Sub-Order List（自動判別）")
        if st.button("🔄 重新載入資料", use_container_width=True):
            st.cache_data.clear()

    st.divider()

    # ── 設定 ──────────────────────────────────────────────────────────────────
    st.markdown("### ⚙️ 設定")
    ref_date   = st.date_input("基準日期（今天）", value=date.today())
    alert_days = st.slider("警示閾值：預估到料 vs 需求差距 ≤ N 天", 7, 90, 30)

    st.divider()
    st.markdown("### 🔍 篩選")
    show_only_alert  = st.checkbox("只顯示警示項目", value=False)
    show_only_future = st.checkbox("只顯示未到料項目", value=True)


# ── load data ────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="讀取 Excel 資料中…")
def load_from_folder(folder: str) -> pd.DataFrame:
    return load_all_files(folder)


def load_from_uploads(files) -> pd.DataFrame:
    """Parse uploaded file objects (BytesIO) using a temp dir."""
    all_records = []
    with tempfile.TemporaryDirectory() as tmp:
        for uf in files:
            dest = Path(tmp) / uf.name
            dest.write_bytes(uf.getvalue())
            try:
                if _is_esp_file(dest):
                    all_records.extend(parse_esp(dest))
                else:
                    all_records.extend(parse_suborder_list(dest))
            except Exception as e:
                st.warning(f"⚠️ 無法解析 {uf.name}：{e}")
    if not all_records:
        return pd.DataFrame()
    df = pd.DataFrame(all_records)
    for col in ("ros", "eta", "ata", "effective_delivery", "sub_order_received"):
        df[col] = pd.to_datetime(df[col], errors="coerce")
    return df


if data_mode == "⬆️ 上傳 Excel 檔案":
    if not uploaded_files:
        st.info("👈 請從左側側邊欄上傳 Excel 報告（ESP 報告或廠商 Sub-Order List）")
        st.stop()
    df_raw = load_from_uploads(uploaded_files)
else:
    df_raw = load_from_folder(data_dir)

if df_raw.empty:
    st.warning("⚠️ 找不到任何可解析的資料，請確認檔案格式後重新上傳。")
    st.stop()


# ── derived columns ──────────────────────────────────────────────────────────
today         = pd.Timestamp(ref_date)
week_start    = today - timedelta(days=today.weekday())
week_end      = week_start + timedelta(days=6)
next_week_end = week_end + timedelta(days=7)

df = df_raw.copy()
df["days_to_delivery"] = (df["effective_delivery"] - today).dt.days
df["days_ros_gap"]     = (df["effective_delivery"] - df["ros"]).dt.days
df["is_arrived"]       = df["ata"].notna()
df["is_this_week"]     = (df["effective_delivery"] >= week_start) & (df["effective_delivery"] <= week_end)
df["is_next_week"]     = (df["effective_delivery"] > week_end)    & (df["effective_delivery"] <= next_week_end)

def alert_level(row):
    if pd.isna(row["effective_delivery"]) or pd.isna(row["ros"]):
        return "unknown"
    if row["is_arrived"]:
        return "arrived"
    gap = row["days_ros_gap"]
    if gap <= 0:
        return "critical"
    if gap <= alert_days:
        return "warning"
    return "ok"

df["alert"] = df.apply(alert_level, axis=1)

if show_only_future:
    df = df[~df["is_arrived"]]
if show_only_alert:
    df = df[df["alert"].isin(["critical", "warning"])]


# ── KPI row ──────────────────────────────────────────────────────────────────
st.markdown("## 📦 材料到料時程警示儀表板")

files_loaded = df_raw["source_file"].nunique()
st.caption(f"已載入 **{files_loaded}** 個檔案 · 共 **{len(df_raw)}** 筆材料項目")

col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("📋 總材料項目",       len(df_raw))
col2.metric("✅ 已到料",           int(df_raw["ata"].notna().sum()))
col3.metric("🔴 緊急警示",         int((df["alert"] == "critical").sum()))
col4.metric("🟠 一個月內到期",      int((df["alert"] == "warning").sum()))
col5.metric("📅 本週 / 次週預計",  f"{int(df['is_this_week'].sum())} / {int(df['is_next_week'].sum())}")

st.divider()


# ── tabs ─────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs(["📅 本週 & 次週到料", "⚠️ 警示清單", "📊 甘特時程圖", "📋 全部明細"])

# ── Tab 1: weekly arrival ─────────────────────────────────────────────────────
with tab1:
    def render_week_section(label: str, mask: pd.Series):
        st.markdown(f'<div class="week-header">{label}</div>', unsafe_allow_html=True)
        rows = df[mask]
        if rows.empty:
            st.info("本期間無預計到料項目")
            return
        for _, r in rows.iterrows():
            alert  = r["alert"]
            cls    = {"critical": "alert-red", "warning": "alert-orange"}.get(alert, "alert-green")
            icon   = "🔴" if alert == "critical" else ("🟠" if alert == "warning" else "🟢")
            eta_str = r["effective_delivery"].strftime("%Y-%m-%d") if pd.notna(r["effective_delivery"]) else "未知"
            ros_str = r["ros"].strftime("%Y-%m-%d") if pd.notna(r["ros"]) else "未定"
            gap_str = f"{int(r['days_ros_gap'])} 天" if pd.notna(r.get("days_ros_gap")) else "—"
            tag_label  = f" `{r['tag_no']}`"   if r.get("tag_no")   else ""
            vendor_tag = f" | 廠商: **{r['sub_vendor']}**" if r["sub_vendor"] else ""
            st.markdown(f"""
<div class="{cls}">
{icon} <b>{r['item_name']}</b>{tag_label}{vendor_tag}<br>
&nbsp;&nbsp;📄 PO: {r['po_no']} &nbsp;|&nbsp; 請購單: {r['mr_no']} &nbsp;|&nbsp; 來源: {r['source_file']}<br>
&nbsp;&nbsp;🚚 預計到料: <b>{eta_str}</b> &nbsp;|&nbsp; 需求時間(ROS): {ros_str} &nbsp;|&nbsp; 差距: {gap_str}
</div>""", unsafe_allow_html=True)

    render_week_section(
        f"🗓️ 本週預計到料  ({week_start.strftime('%m/%d')} – {week_end.strftime('%m/%d')})",
        df["is_this_week"],
    )
    st.markdown("")
    render_week_section(
        f"🗓️ 次週預計到料  ({(week_end + timedelta(1)).strftime('%m/%d')} – {next_week_end.strftime('%m/%d')})",
        df["is_next_week"],
    )

# ── Tab 2: alert list ─────────────────────────────────────────────────────────
with tab2:
    st.subheader("⚠️ 警示項目清單")
    alert_df = df[df["alert"].isin(["critical", "warning"])].copy()
    if alert_df.empty:
        st.success("🎉 目前無需警示的材料項目！")
    else:
        alert_df["警示等級"] = alert_df["alert"].map({"critical": "🔴 緊急", "warning": "🟠 注意"})
        alert_df["預計到料"] = alert_df["effective_delivery"].dt.strftime("%Y-%m-%d")
        alert_df["需求時間"] = alert_df["ros"].dt.strftime("%Y-%m-%d")
        alert_df["差距(天)"] = alert_df["days_ros_gap"].apply(lambda x: int(x) if pd.notna(x) else None)

        display_cols = {
            "警示等級": "警示", "mr_no": "請購單號", "po_no": "PO編號",
            "tag_no": "設備Tag", "item_name": "材料名稱",
            "sub_vendor": "次廠商", "vendor": "主廠商",
            "預計到料": "預計到料", "需求時間": "需求時間(ROS)",
            "差距(天)": "差距(天)", "source_file": "來源檔案",
        }
        out = alert_df[list(display_cols.keys())].rename(columns=display_cols)
        st.dataframe(out, use_container_width=True, hide_index=True,
                     column_config={"差距(天)": st.column_config.NumberColumn(format="%d 天")})
        st.download_button("⬇️ 下載警示清單 CSV",
                           out.to_csv(index=False).encode("utf-8-sig"),
                           "alert_list.csv", "text/csv")

# ── Tab 3: Gantt ─────────────────────────────────────────────────────────────
with tab3:
    st.subheader("📊 材料到料甘特時程圖")
    gantt_df = df[df["effective_delivery"].notna()].copy()

    c1, c2 = st.columns(2)
    with c1:
        po_options = ["全部"] + sorted(gantt_df["po_no"].dropna().unique().tolist())
        selected_po = st.selectbox("篩選 PO 編號", po_options)
    with c2:
        weeks_ahead = st.slider("顯示未來幾週", 2, 52, 12)

    if selected_po != "全部":
        gantt_df = gantt_df[gantt_df["po_no"] == selected_po]
    gantt_df = gantt_df[gantt_df["effective_delivery"] <= today + timedelta(weeks=weeks_ahead)]

    color_map = {"緊急": "#d32f2f", "注意": "#f57c00", "正常": "#388e3c",
                 "已到料": "#1565c0", "未知": "#9e9e9e"}

    if gantt_df.empty:
        st.info("所選條件下無資料")
    else:
        gantt_df["Start"]  = gantt_df["effective_delivery"] - timedelta(days=1)
        gantt_df["Finish"] = gantt_df["effective_delivery"]
        gantt_df["Color"]  = gantt_df["alert"].map(
            {"critical": "緊急", "warning": "注意", "ok": "正常",
             "arrived": "已到料", "unknown": "未知"})
        gantt_df["label"]  = gantt_df.apply(
            lambda r: f"{r['item_name']}" + (f" [{r['tag_no']}]" if r.get("tag_no") else ""), axis=1)

        fig = px.timeline(
            gantt_df, x_start="Start", x_end="Finish", y="label", color="Color",
            color_discrete_map=color_map,
            hover_data={"mr_no": True, "sub_vendor": True, "ros": True, "effective_delivery": True},
            labels={"label": "材料", "Color": "狀態"},
            height=max(400, len(gantt_df) * 28 + 100),
        )
        fig.add_vline(x=today, line_color="blue", line_dash="dash", annotation_text="今天")
        fig.update_layout(xaxis_title="日期", yaxis_title="", legend_title="狀態",
                          margin=dict(l=20, r=20, t=40, b=20))
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("📈 到料時程散佈圖")
    scatter_df = df[df["effective_delivery"].notna() & df["ros"].notna()].copy()
    if not scatter_df.empty:
        scatter_df["Color"] = scatter_df["alert"].map(
            {"critical": "緊急", "warning": "注意", "ok": "正常",
             "arrived": "已到料", "unknown": "未知"})
        fig2 = px.scatter(
            scatter_df, x="days_to_delivery", y="days_ros_gap", color="Color",
            color_discrete_map=color_map,
            hover_data=["item_name", "mr_no", "po_no", "sub_vendor"],
            labels={"days_to_delivery": "距今天數 (天)",
                    "days_ros_gap": "預計到料 vs ROS 差距 (天，負=遲到)",
                    "Color": "狀態"},
        )
        fig2.add_hline(y=0, line_color="red", line_dash="dash", annotation_text="ROS")
        fig2.add_hline(y=alert_days, line_color="orange", line_dash="dot",
                       annotation_text=f"警示閾值 ({alert_days}天)")
        st.plotly_chart(fig2, use_container_width=True)

# ── Tab 4: full table ─────────────────────────────────────────────────────────
with tab4:
    st.subheader("📋 全部材料明細")
    search = st.text_input("搜尋 (材料名稱 / PO / 廠商)", "")
    view_df = df.copy()
    if search:
        mask = (
            view_df["item_name"].str.contains(search, case=False, na=False) |
            view_df["po_no"].str.contains(search, case=False, na=False) |
            view_df["sub_vendor"].str.contains(search, case=False, na=False) |
            view_df["vendor"].str.contains(search, case=False, na=False)
        )
        view_df = view_df[mask]

    view_df["警示"] = view_df["alert"].map({
        "critical": "🔴 緊急", "warning": "🟠 注意",
        "ok": "🟢 正常", "arrived": "✅ 已到料", "unknown": "❓ 未知"
    })
    out_cols = {
        "警示": "警示", "source_file": "來源檔案", "mr_no": "請購單號",
        "po_no": "PO編號", "tag_no": "設備Tag", "item_name": "材料名稱",
        "qty": "數量", "vendor": "主廠商", "sub_vendor": "次廠商",
        "eta": "預計到料(ETA)", "ata": "實際到料(ATA)",
        "ros": "需求時間(ROS)", "days_ros_gap": "差距(天)",
        "sheet": "工作表", "record_type": "類型",
    }
    out = view_df[list(out_cols.keys())].rename(columns=out_cols)
    st.dataframe(out, use_container_width=True, hide_index=True)
    st.download_button("⬇️ 下載完整清單 CSV",
                       out.to_csv(index=False).encode("utf-8-sig"),
                       "full_material_list.csv", "text/csv")
