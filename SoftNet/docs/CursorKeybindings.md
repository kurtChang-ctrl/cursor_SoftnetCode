# Cursor 折疊快捷鍵設定指南

## 問題說明
工作區的 `.vscode/keybindings.json` 可能不會自動生效，需要在 Cursor 的使用者層級設定快捷鍵。

## 方法 1：透過 Cursor UI 設定（推薦）

### 步驟：
1. **開啟快捷鍵設定**
   - 按 `Ctrl + K, Ctrl + S`（或 `Ctrl + Shift + P` → 輸入 "Preferences: Open Keyboard Shortcuts"）

2. **搜尋 "Fold All"**
   - 在搜尋框輸入：`fold all`
   - 找到 "Editor: Fold All" 命令

3. **設定快捷鍵**
   - 點擊 "Fold All" 左側的「+」圖示（或雙擊該項目）
   - 按下您想要的快捷鍵：先按 `Ctrl + K`，然後按 `Ctrl + 0`
   - 如果出現衝突提示，選擇「Replace」替換現有快捷鍵

4. **同樣設定其他折疊命令**
   - **Unfold All**: `Ctrl + K, Ctrl + J`
   - **Fold Level 1**: `Ctrl + K, Ctrl + 1`
   - **Fold Level 2**: `Ctrl + K, Ctrl + 2`
   - **Fold Level 3**: `Ctrl + K, Ctrl + 3`

## 方法 2：直接編輯使用者設定檔

### 步驟：
1. **找到 Cursor 的使用者設定目錄**
   - Windows 路徑：`%APPDATA%\Cursor\User\keybindings.json`
   - 完整路徑通常是：`C:\Users\您的使用者名稱\AppData\Roaming\Cursor\User\keybindings.json`

2. **開啟 keybindings.json**
   - 按 `Ctrl + Shift + P` → 輸入 "Preferences: Open User Keyboard Shortcuts (JSON)"
   - 或直接開啟上述路徑的檔案

3. **加入以下設定**
   ```json
   [
     {
       "key": "ctrl+k ctrl+0",
       "command": "editor.foldAll",
       "when": "editorTextFocus"
     },
     {
       "key": "ctrl+k ctrl+j",
       "command": "editor.unfoldAll",
       "when": "editorTextFocus"
     },
     {
       "key": "ctrl+k ctrl+1",
       "command": "editor.foldLevel1",
       "when": "editorTextFocus"
     },
     {
       "key": "ctrl+k ctrl+2",
       "command": "editor.foldLevel2",
       "when": "editorTextFocus"
     },
     {
       "key": "ctrl+k ctrl+3",
       "command": "editor.foldLevel3",
       "when": "editorTextFocus"
     }
   ]
   ```

4. **儲存檔案並重新載入 Cursor**
   - 按 `Ctrl + Shift + P` → 輸入 "Reload Window"

## 方法 3：使用 PowerShell 腳本自動設定

執行專案中的腳本：
```powershell
.\scripts\Setup-Cursor-Keybindings.ps1
```

## 測試快捷鍵

設定完成後，在任何程式碼檔案中測試：
- `Ctrl + K, Ctrl + 0` - 應該會折疊所有區塊
- `Ctrl + K, Ctrl + J` - 應該會展開所有區塊
- `Ctrl + K, Ctrl + 1` - 應該會折疊到第 1 層級

## 如果還是不行

1. **檢查是否有快捷鍵衝突**
   - 在快捷鍵設定中搜尋 `ctrl+k ctrl+0`
   - 查看是否有其他命令使用相同快捷鍵

2. **嘗試其他快捷鍵組合**
   - `Ctrl + Shift + [` - 折疊當前區塊
   - `Ctrl + Shift + ]` - 展開當前區塊

3. **使用命令面板**
   - `Ctrl + Shift + P` → 輸入 "Fold All" 或 "Unfold All"

## 常用折疊快捷鍵總覽

| 功能 | 快捷鍵 | 命令 |
|------|--------|------|
| 折疊所有 | `Ctrl+K, Ctrl+0` | `editor.foldAll` |
| 展開所有 | `Ctrl+K, Ctrl+J` | `editor.unfoldAll` |
| 折疊第 1 層 | `Ctrl+K, Ctrl+1` | `editor.foldLevel1` |
| 折疊第 2 層 | `Ctrl+K, Ctrl+2` | `editor.foldLevel2` |
| 折疊第 3 層 | `Ctrl+K, Ctrl+3` | `editor.foldLevel3` |
| 折疊當前區塊 | `Ctrl+Shift+[` | `editor.fold` |
| 展開當前區塊 | `Ctrl+Shift+]` | `editor.unfold` |

