# 📬 自動寄信程式（Auto Email Sender）
這是一個使用 Python 撰寫的自動寄信工具，支援收件人名單讀取、附件添加、排程發送等功能。適合用於日常報表通知、自動提醒等場景。

# 📌 功能特色
✅ 支援 Excel 或 TXT 檔作為收件人名單來源

✅ 可添加多個附件寄送

✅ 支援立即發送或每日定時寄信

✅ 記錄每次發送結果與錯誤於 email_sender.log

# 🔧 使用技術
smtplib, email.mime：寄送 Email
schedule：設定每日定時任務
pandas, openpyxl：處理收件人 Excel 名單
logging：日誌記錄發送結果
os, sys, datetime：檔案與時間處理

🧠 學習筆記
這個專案是為了學習如何用 Python 完成實用的自動化應用。過程中練習了：
發送郵件流程與 MIME 格式
檔案處理、附件打包技巧
如何設計使用者互動流程（CLI）
定時任務與日誌管理
