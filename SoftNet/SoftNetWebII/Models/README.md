# ONNX 模型目錄

此目錄用於存放 ONNX embedding 模型文件。

## 快速開始

### 方法 1：使用 PowerShell 腳本（推薦）

在 PowerShell 中執行：

```powershell
cd SoftNetWebII\Models
.\download-onnx-model.ps1
```

腳本會自動下載模型並放置到正確位置。

### 方法 2：手動下載

1. 前往 Hugging Face：https://huggingface.co/sentence-transformers/all-MiniLM-L6-v2
2. 下載以下檔案：
   - **ONNX 模型**：https://huggingface.co/sentence-transformers/all-MiniLM-L6-v2/resolve/main/onnx/model.onnx
     - 重新命名為 `all-MiniLM-L6-v2.onnx`
   - **詞彙表**：https://huggingface.co/sentence-transformers/all-MiniLM-L6-v2/resolve/main/vocab.txt
3. 將兩個檔案都放置在此目錄

## 模型資訊

- **模型名稱**：all-MiniLM-L6-v2
- **輸出維度**：384
- **最大序列長度**：512 tokens
- **模型大小**：約 80-90 MB
- **vocab.txt 大小**：約 226 KB
- **Tokenizer**：完整的 BERT tokenizer（支援 WordPiece tokenization）
  - 使用單例模式快取，避免重複載入
  - 支援多語言（中文、日文、韓文等）
  - 正確處理特殊 token（[CLS], [SEP], [UNK], [PAD]）

## 編譯設定

模型檔案和 vocab.txt 會在編譯時自動複製到輸出目錄（`bin\Debug\net10.0\Models\` 或 `bin\Release\net10.0\Models\`）。

**檔案清單：**
- `all-MiniLM-L6-v2.onnx` - ONNX 模型檔案
- `vocab.txt` - BERT 詞彙表（用於完整的 tokenizer）

如果模型檔案不存在，系統會自動回退到 OpenAI API。  
如果 vocab.txt 不存在，系統會使用簡化的 tokenizer（準確度較低）。

## 注意事項

1. **模型文件較大**（約 80-90 MB），請確保有足夠的磁碟空間
2. **vocab.txt 檔案**（約 226 KB）用於完整的 BERT tokenizer，建議一併下載
3. **首次使用時**，系統會自動載入模型和 tokenizer（約需 1-2 秒）
4. **Tokenizer 快取**：使用單例模式，避免重複載入，提升效能
5. **如果模型文件不存在**，系統會自動回退到 OpenAI API（需要設定 `OpenAIApiKey`）
6. **如果 vocab.txt 不存在**，系統會使用簡化的 tokenizer（準確度較低）
7. **編譯後**，所有檔案會自動複製到輸出目錄，無需手動複製

## 驗證設定

啟動應用程式後，檢查日誌：

- **ONNX 模型成功**：`ONNX 模型載入成功: C:\...\bin\Debug\net10.0\Models\all-MiniLM-L6-v2.onnx`
- **BERT Tokenizer 成功**：`BERT Tokenizer 載入成功: C:\...\bin\Debug\net10.0\Models\vocab.txt（詞彙表大小: 30522）`
- **使用完整 tokenizer**：`使用 ONNX 模型生成 embedding（維度: 384）`
- **使用簡化 tokenizer（警告）**：`vocab.txt 不存在: ...，將使用簡化 tokenizer`
- **模型不存在（自動回退）**：`ONNX 模型文件不存在: ...，將使用 OpenAI API 作為備用方案`

## 替代方案

如果無法下載 ONNX 模型，可以：
1. 使用 OpenAI API（需要設定 `OpenAIApiKey`）
2. 使用其他 embedding 服務
