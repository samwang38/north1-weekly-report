# 北一區週報產生器

## 首次安裝（新同事）

在終端機貼上以下指令，一鍵完成：

```bash
curl -fsSL https://raw.githubusercontent.com/samwang38/north1-weekly-report/main/install.sh | bash
```

安裝完成後，桌面會出現 `北一區週報-app` 資料夾，雙擊其中的 `啟動北一區週報.command` 即可使用。

---

## 日常使用

1. **連上公司 VPN**（需能連至 192.168.1.177）
2. 雙擊 `啟動北一區週報.command`
   - 每次啟動會自動同步最新版本
   - 第一次執行時會自動安裝必要 Python 套件
3. 瀏覽器開啟：`http://127.0.0.1:8782/`
4. 選擇週結束日期（週六），點「產生週報」
5. 等候約 2 分鐘後，點「下載 Excel」

---

## 環境需求

| 項目 | 需求 |
|------|------|
| Python | 3.8 以上 |
| JDK | 1.8（路徑：`/Library/Java/JavaVirtualMachines/jdk1.8.0_251.jdk`）|
| EPBrowser lib | `/Library/EPBrowser/EPB/Shell/` |
| 網路 | 公司 VPN（192.168.1.177:8080）|
| OS | macOS |

## 資料夾結構

```
北一區週報-app/
├── server.py               主程式（HTTP server + 填表邏輯）
├── multistore_engine.py    EPB 資料查詢與計算引擎
├── EPBReportQuery.java     Java EPB 橋接器原始碼
├── 啟動北一區週報.command   雙擊啟動腳本（含自動更新）
├── install.sh              一鍵安裝腳本
├── requirements.txt        Python 套件清單
├── data/
│   └── SAcare對應價目表.xlsx
├── template/
│   └── 北一區週報_優化.xlsx  週報範本
└── static/                 前端網頁
    ├── index.html
    ├── styles.css
    └── app.js
```

## 注意事項

- 關閉終端機視窗即可停止伺服器
- 同一時間只支援一個產生工作
- 若出現 `EPB 連線失敗`，請確認 VPN 已連線
