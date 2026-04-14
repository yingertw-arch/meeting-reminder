# 會議提醒管理 Meeting Reminder

Windows 桌面會議提醒工具，支援單次 / 每天 / 每週 / 每月週期提醒，時間到會彈出全螢幕誇張警報。

## 功能
- 多筆提醒清單，依時間自動排序
- 週期設定：單次 / 每天 / 週一到五 / 每週指定日 / 每月指定日
- 時間到彈出全螢幕警報（含震動、閃爍效果）
- 縮小至系統匣背景監控
- 開機自動啟動
- 啟動時自動清理過期提醒

## 安裝方式

### 方法一：直接下載安裝程式
前往 [Releases](../../releases) 下載最新版 `會議提醒_安裝程式.exe`，雙擊安裝即可。

### 方法二：從原始碼執行
需要 Python 3.8+

```bash
pip install pystray Pillow
python meeting_reminder.py
```

## 打包方式
```bash
build.bat
```
接著用 Inno Setup 開啟 `installer.iss` 並按 Ctrl+F9 編譯安裝程式。
