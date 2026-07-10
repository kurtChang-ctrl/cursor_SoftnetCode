using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace SoftNetWebII.Services
{
    /// <summary>
    /// BERT Tokenizer 實作
    /// 支援 WordPiece tokenization 和 vocab.txt
    /// 使用單例模式快取實例，避免重複載入
    /// </summary>
    public class BertTokenizer
    {
        private static BertTokenizer _instance = null;
        private static readonly object _lock = new object();

        private readonly Dictionary<string, int> _vocab;
        private readonly Dictionary<int, string> _idsToTokens;
        private const string UNK_TOKEN = "[UNK]";
        private const string CLS_TOKEN = "[CLS]";
        private const string SEP_TOKEN = "[SEP]";
        private const string PAD_TOKEN = "[PAD]";
        private const string MASK_TOKEN = "[MASK]";

        public int UNK_TOKEN_ID => _vocab[UNK_TOKEN];
        public int CLS_TOKEN_ID => _vocab[CLS_TOKEN];
        public int SEP_TOKEN_ID => _vocab[SEP_TOKEN];
        public int PAD_TOKEN_ID => _vocab[PAD_TOKEN];
        public int MASK_TOKEN_ID => _vocab[MASK_TOKEN];

        /// <summary>
        /// 取得單例實例（快取 tokenizer，避免重複載入）
        /// </summary>
        public static BertTokenizer GetInstance(string vocabPath)
        {
            if (_instance == null)
            {
                lock (_lock)
                {
                    if (_instance == null)
                    {
                        _instance = new BertTokenizer(vocabPath);
                    }
                }
            }
            return _instance;
        }

        /// <summary>
        /// 從 vocab.txt 載入詞彙表
        /// </summary>
        private BertTokenizer(string vocabPath)
        {
            if (!File.Exists(vocabPath))
            {
                throw new FileNotFoundException($"vocab.txt 檔案不存在: {vocabPath}");
            }

            _vocab = new Dictionary<string, int>();
            _idsToTokens = new Dictionary<int, string>();

            var lines = File.ReadAllLines(vocabPath, Encoding.UTF8);
            for (int i = 0; i < lines.Length; i++)
            {
                var token = lines[i].Trim();
                if (!string.IsNullOrEmpty(token))
                {
                    _vocab[token] = i;
                    _idsToTokens[i] = token;
                }
            }
        }

        /// <summary>
        /// 將文字轉換為 token IDs
        /// </summary>
        /// <param name="text">輸入文字</param>
        /// <param name="maxLength">最大長度（預設 512）</param>
        /// <param name="addSpecialTokens">是否加入特殊 token（[CLS], [SEP]）</param>
        /// <returns>Token IDs 陣列</returns>
        public long[] Encode(string text, int maxLength = 512, bool addSpecialTokens = true)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                // 如果文字為空，返回只有特殊 token 的陣列
                var emptyTokens = new long[maxLength];
                if (addSpecialTokens)
                {
                    emptyTokens[0] = CLS_TOKEN_ID;
                    emptyTokens[1] = SEP_TOKEN_ID;
                    for (int i = 2; i < maxLength; i++)
                    {
                        emptyTokens[i] = PAD_TOKEN_ID;
                    }
                }
                return emptyTokens;
            }

            // 1. 基本清理和正規化（轉小寫，但保留中文等）
            text = NormalizeText(text);
            
            // 2. 基本分詞（按空格和標點符號）
            var words = BasicTokenize(text);
            
            // 3. WordPiece tokenization
            var tokens = new List<string>();
            
            if (addSpecialTokens)
            {
                tokens.Add(CLS_TOKEN);
            }

            foreach (var word in words)
            {
                var wordTokens = WordPieceTokenize(word);
                tokens.AddRange(wordTokens);
            }

            if (addSpecialTokens)
            {
                tokens.Add(SEP_TOKEN);
            }

            // 4. 轉換為 token IDs
            var tokenIds = new List<long>();
            foreach (var token in tokens)
            {
                if (_vocab.TryGetValue(token, out int tokenId))
                {
                    tokenIds.Add(tokenId);
                }
                else
                {
                    // 如果找不到 token，使用 [UNK]
                    tokenIds.Add(UNK_TOKEN_ID);
                }
            }

            // 5. 截斷或填充到 maxLength
            if (tokenIds.Count > maxLength)
            {
                tokenIds = tokenIds.Take(maxLength).ToList();
                // 確保最後一個是 SEP（如果啟用特殊 token）
                if (addSpecialTokens && tokenIds.Count > 0 && tokenIds[tokenIds.Count - 1] != SEP_TOKEN_ID)
                {
                    tokenIds[tokenIds.Count - 1] = SEP_TOKEN_ID;
                }
            }
            else if (tokenIds.Count < maxLength)
            {
                // 填充 PAD
                while (tokenIds.Count < maxLength)
                {
                    tokenIds.Add(PAD_TOKEN_ID);
                }
            }

            return tokenIds.ToArray();
        }

        /// <summary>
        /// 文字正規化（轉小寫，但保留中文、日文等）
        /// </summary>
        private string NormalizeText(string text)
        {
            var sb = new StringBuilder();
            foreach (char c in text)
            {
                // 如果是英文字母，轉小寫
                if (c >= 'A' && c <= 'Z')
                {
                    sb.Append((char)(c + 32));
                }
                // 保留中文、日文、韓文等（Unicode 範圍）
                else if (IsCJKCharacter(c))
                {
                    sb.Append(c);
                }
                // 保留其他字元
                else
                {
                    sb.Append(c);
                }
            }
            return sb.ToString().Trim();
        }

        /// <summary>
        /// 判斷是否為中日韓字元
        /// </summary>
        private bool IsCJKCharacter(char c)
        {
            // 中文、日文、韓文 Unicode 範圍
            return (c >= 0x4E00 && c <= 0x9FFF) ||   // CJK Unified Ideographs
                   (c >= 0x3040 && c <= 0x309F) ||   // Hiragana
                   (c >= 0x30A0 && c <= 0x30FF) ||   // Katakana
                   (c >= 0xAC00 && c <= 0xD7AF);      // Hangul
        }

        /// <summary>
        /// 基本分詞（按空格、標點符號等）
        /// </summary>
        private List<string> BasicTokenize(string text)
        {
            var words = new List<string>();
            var currentWord = new StringBuilder();

            foreach (char c in text)
            {
                if (char.IsWhiteSpace(c))
                {
                    if (currentWord.Length > 0)
                    {
                        words.Add(currentWord.ToString());
                        currentWord.Clear();
                    }
                }
                else if (IsPunctuation(c))
                {
                    if (currentWord.Length > 0)
                    {
                        words.Add(currentWord.ToString());
                        currentWord.Clear();
                    }
                    words.Add(c.ToString());
                }
                else
                {
                    currentWord.Append(c);
                }
            }

            if (currentWord.Length > 0)
            {
                words.Add(currentWord.ToString());
            }

            return words;
        }

        /// <summary>
        /// WordPiece tokenization
        /// 將單詞分解為子詞（subwords）
        /// </summary>
        private List<string> WordPieceTokenize(string word)
        {
            var outputTokens = new List<string>();

            // 如果單詞在 vocab 中，直接返回
            if (_vocab.ContainsKey(word))
            {
                outputTokens.Add(word);
                return outputTokens;
            }

            // 嘗試將單詞分解為子詞
            int start = 0;
            var subTokens = new List<string>();

            while (start < word.Length)
            {
                int end = word.Length;
                string curSubstr = null;

                // 從最長的子詞開始嘗試
                while (start < end)
                {
                    string substr = word.Substring(start, end - start);
                    
                    // 如果不是第一個子詞，加上 ## 前綴
                    if (start > 0)
                    {
                        substr = "##" + substr;
                    }

                    if (_vocab.ContainsKey(substr))
                    {
                        curSubstr = substr;
                        break;
                    }

                    end--;
                }

                // 如果找不到任何子詞，使用 [UNK]
                if (curSubstr == null)
                {
                    subTokens.Add(UNK_TOKEN);
                    break;
                }

                subTokens.Add(curSubstr);
                start = end;
            }

            return subTokens;
        }

        /// <summary>
        /// 判斷是否為標點符號
        /// </summary>
        private bool IsPunctuation(char c)
        {
            return char.IsPunctuation(c) || 
                   c == '!' || c == '?' || c == '.' || c == ',' || 
                   c == ';' || c == ':' || c == '-' || c == '_' ||
                   c == '(' || c == ')' || c == '[' || c == ']' ||
                   c == '{' || c == '}' || c == '"' || c == '\'' ||
                   c == '/' || c == '\\' || c == '@' || c == '#' ||
                   c == '$' || c == '%' || c == '^' || c == '&' ||
                   c == '*' || c == '+' || c == '=' || c == '|' ||
                   c == '~' || c == '`' || c == '<' || c == '>';
        }

        /// <summary>
        /// 將 token IDs 轉換回文字（用於除錯）
        /// </summary>
        public string Decode(long[] tokenIds)
        {
            var tokens = new List<string>();
            foreach (var id in tokenIds)
            {
                if (_idsToTokens.TryGetValue((int)id, out string token))
                {
                    tokens.Add(token);
                }
            }
            return string.Join(" ", tokens);
        }

        /// <summary>
        /// 取得詞彙表大小
        /// </summary>
        public int VocabSize => _vocab.Count;
    }
}
