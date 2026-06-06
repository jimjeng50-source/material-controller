# 材料到料時程警示儀表板

## 快速啟動

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 使用方式

1. 將所有 ESP 報告 (.xlsx) 和廠商 Sub-Order List 放入 `data/` 資料夾
2. 執行 `streamlit run app.py`
3. 左側邊欄可指定資料夾路徑、基準日期、警示天數閾值

## 功能

- **自動判別** ESP 報告格式 vs 廠商 Sub-Order List 格式
- **本週 / 次週** 預計到料材料一覽（含警示色標）
- **警示清單**：預計到料 vs 需求時間差距 ≤ 設定天數（預設 30 天）
- **甘特圖**：可依 PO 篩選、顯示未來 N 週時程
- **散佈圖**：距今天數 vs ROS 差距，一眼看出風險分佈
- **全部明細**：搜尋、下載 CSV

## 資料夾結構

```
material-controller/
├── app.py              # Streamlit 主程式
├── src/
│   └── parser.py       # Excel 解析器（ESP + Sub-Order）
├── data/               # 放入所有 Excel 報告
└── requirements.txt
```
