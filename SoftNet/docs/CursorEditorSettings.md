# Cursor 編輯器設定指南

## 查看目前行距設定

### 方法 1：透過設定 UI
1. 按 `Ctrl + ,` 開啟設定
2. 在搜尋框輸入：`line height`
3. 查看 "Editor: Line Height" 的數值

### 方法 2：查看設定檔
1. 按 `Ctrl + Shift + P` → 輸入 "Preferences: Open User Settings (JSON)"
2. 查看是否有 `"editor.lineHeight"` 設定

### 方法 3：使用命令列
```powershell
Get-Content "$env:APPDATA\Cursor\User\settings.json" | Select-String "lineHeight"
```

## 行距設定說明

### 預設值
- 如果沒有設定 `editor.lineHeight`，預設值為 `0`（自動計算）
- 自動計算時，行距約為字型大小的 **1.2-1.5 倍**

### 設定格式
```json
{
  "editor.lineHeight": 24
}
```

### 數值說明
- **0** - 自動計算（預設，約為字型大小的 1.2-1.5 倍）
- **數字** - 固定像素值（例如：`24` 表示 24 像素）
- **小數** - 相對於字型大小的倍數（例如：`1.5` 表示字型大小的 1.5 倍）

### 常用行距值
| 字型大小 | 建議行距（像素） | 建議行距（倍數） |
|---------|----------------|----------------|
| 12px | 18-20 | 1.5-1.67 |
| 14px | 20-22 | 1.43-1.57 |
| 16px | 24-26 | 1.5-1.625 |
| 18px | 27-30 | 1.5-1.67 |
| 20px | 30-33 | 1.5-1.65 |

## 設定行距

### 方法 1：透過設定 UI（推薦）
1. 按 `Ctrl + ,` 開啟設定
2. 搜尋 "line height"
3. 點擊 "Editor: Line Height"
4. 輸入數值（例如：`24` 或 `1.5`）
5. 設定會自動儲存

### 方法 2：直接編輯設定檔
1. 按 `Ctrl + Shift + P` → 輸入 "Preferences: Open User Settings (JSON)"
2. 加入或修改：
   ```json
   {
     "editor.lineHeight": 24
   }
   ```
3. 儲存檔案

### 方法 3：使用 PowerShell 腳本
```powershell
# 設定行距為 24 像素
$settingsFile = "$env:APPDATA\Cursor\User\settings.json"
$settings = Get-Content $settingsFile -Raw | ConvertFrom-Json
$settings | Add-Member -MemberType NoteProperty -Name "editor.lineHeight" -Value 24 -Force
$settings | ConvertTo-Json -Depth 10 | Set-Content $settingsFile
```

## 其他相關編輯器設定

### 字型大小
```json
{
  "editor.fontSize": 14
}
```

### 字型家族
```json
{
  "editor.fontFamily": "Consolas, 'Courier New', monospace"
}
```

### 字型粗細
```json
{
  "editor.fontWeight": "normal"
}
```

### 字母間距
```json
{
  "editor.letterSpacing": 0.5
}
```

## 快速查看目前所有編輯器設定

執行以下 PowerShell 命令：
```powershell
Get-Content "$env:APPDATA\Cursor\User\settings.json" | ConvertFrom-Json | Select-Object -Property editor.*
```

## 建議設定（適合程式開發）

```json
{
  "editor.fontSize": 14,
  "editor.lineHeight": 22,
  "editor.fontFamily": "Consolas, 'Courier New', monospace",
  "editor.fontWeight": "normal",
  "editor.letterSpacing": 0.5
}
```

這些設定提供：
- 清晰易讀的字型
- 適當的行距（不會太緊或太鬆）
- 良好的程式碼可讀性

