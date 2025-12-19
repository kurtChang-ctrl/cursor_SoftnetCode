# GitHub 同步指南

## 目前狀況分析

### ✅ 應該提交到 GitHub 的檔案

這些檔案是專案配置和文檔，應該與團隊共享：

1. **`.cursor/` 目錄**
   - `.cursor/rules/auto-context.mdc` - Cursor 自動上下文規則
   - 用途：讓團隊成員都能使用相同的 Cursor 上下文設定

2. **`.cursorrules`**
   - Cursor 專案規則檔案
   - 用途：定義專案規範和上下文提示

3. **`.cursorignore`**（已修改）
   - Cursor 忽略檔案規則
   - 用途：排除不需要索引的檔案

4. **`docs/` 目錄**
   - `docs/CursorAutoContext.md` - Cursor 自動上下文說明
   - `docs/CursorKeybindings.md` - 快捷鍵設定指南
   - `docs/CursorEditorSettings.md` - 編輯器設定指南
   - 用途：團隊文檔和設定說明

5. **`scripts/` 目錄**
   - `scripts/Start-Cursor-SoftNet.ps1` - 啟動 Cursor 腳本
   - `scripts/Start-Cursor-SoftNet.bat` - 啟動 Cursor 腳本
   - `scripts/Setup-Cursor-Keybindings.ps1` - 快捷鍵設定腳本
   - 用途：團隊共用的工具腳本

6. **`ProjectRules.md`**
   - 專案規範文件
   - 用途：專案開發規範

### ❌ 不應該提交到 GitHub 的檔案

這些是個人偏好設定，不應該提交：

1. **`.vscode/keybindings.json`**
   - 快捷鍵設定（個人偏好）
   - 已被 `.gitignore` 排除（正確）

2. **使用者層級的 Cursor 設定**
   - `%APPDATA%\Cursor\User\settings.json` - 個人編輯器設定（行距等）
   - `%APPDATA%\Cursor\User\keybindings.json` - 個人快捷鍵設定
   - 這些不在專案目錄中，不會被提交（正確）

## 建議的提交步驟

### 1. 檢查要提交的檔案
```powershell
git status .cursor .cursorrules .cursorignore docs scripts ProjectRules.md
```

### 2. 加入應該提交的檔案
```powershell
git add .cursor/
git add .cursorrules
git add .cursorignore
git add docs/
git add scripts/
git add ProjectRules.md
```

### 3. 確認沒有加入不應該提交的檔案
```powershell
# 確認 .vscode/keybindings.json 不會被提交
git check-ignore .vscode/keybindings.json
# 應該顯示：.vscode/keybindings.json
```

### 4. 提交變更
```powershell
git commit -m "feat(config): 新增 Cursor 設定和專案文檔

- 新增 .cursorrules 和 .cursor/rules 自動上下文設定
- 新增 Cursor 使用說明文檔
- 新增專案規範文件 ProjectRules.md
- 更新 .cursorignore 排除規則"
```

### 5. 推送到 GitHub
```powershell
git push origin main
```

## 檔案說明

### `.cursorignore` vs `.gitignore`

- **`.gitignore`** - Git 版本控制忽略規則
  - 決定哪些檔案不會被 Git 追蹤
  - 例如：`.vscode/` 目錄（個人設定）

- **`.cursorignore`** - Cursor IDE 索引忽略規則
  - 決定哪些檔案不會被 Cursor 索引（用於 AI 上下文）
  - 例如：`bin/`, `obj/`, `.vs/` 等編譯產物

### 為什麼 `.vscode/keybindings.json` 不應該提交？

- 快捷鍵是個人偏好設定
- 每個開發者可能有不同的快捷鍵習慣
- 工作區的 `keybindings.json` 會覆蓋使用者設定，造成困擾
- 因此應該保留在使用者層級（`%APPDATA%\Cursor\User\keybindings.json`）

### 為什麼 `.cursor/` 和 `.cursorrules` 應該提交？

- 這些是專案層級的配置
- 幫助團隊成員使用相同的 Cursor 上下文設定
- 確保 AI 助手對專案有相同的理解

## 驗證清單

提交前請確認：

- [ ] `.cursor/` 目錄已加入
- [ ] `.cursorrules` 已加入
- [ ] `.cursorignore` 已加入
- [ ] `docs/` 目錄已加入
- [ ] `scripts/` 目錄已加入
- [ ] `ProjectRules.md` 已加入
- [ ] `.vscode/keybindings.json` **沒有**被加入（被 gitignore 排除）
- [ ] 沒有其他個人設定檔案被意外加入

## 常見問題

### Q: 為什麼我的快捷鍵設定沒有同步到 GitHub？
A: 這是正確的！快捷鍵是個人偏好，應該保留在使用者層級設定中。每個開發者可以根據自己的習慣設定。

### Q: 團隊成員如何獲得相同的 Cursor 設定？
A: 
1. 拉取專案後，`.cursorrules` 和 `.cursor/rules/` 會自動生效
2. 個人快捷鍵設定需要執行 `scripts/Setup-Cursor-Keybindings.ps1`
3. 個人編輯器設定（行距等）需要手動設定或參考 `docs/CursorEditorSettings.md`

### Q: 如果我想分享我的快捷鍵設定給團隊怎麼辦？
A: 可以將快捷鍵設定加入 `docs/CursorKeybindings.md` 作為參考，但不要提交到 `.vscode/keybindings.json`。
