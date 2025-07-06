using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq; // Add this line
using System.Text.RegularExpressions;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SlideSCI
{
    public class CodeHighlighter
    {
        private Dictionary<string, Color> themeColors;
        private Dictionary<
            string,
            List<(string pattern, RegexOptions options, string type)>
        > languagePatterns;
        private Dictionary<string, string> languageAliases;
        private HashSet<string> processedRanges;

        public CodeHighlighter(bool isDarkTheme)
        {
            processedRanges = new HashSet<string>();
            InitializeColors(isDarkTheme);
            InitializePatterns();
            InitializeAliases();
        }

        private void InitializeColors(bool isDarkTheme)
        {
            if (isDarkTheme)
            {
                themeColors = new Dictionary<string, Color>
                {
                    { "keyword", Color.FromArgb(86, 156, 214) }, // 蓝色
                    { "comment", Color.FromArgb(87, 166, 74) }, // 绿色
                    { "string", Color.FromArgb(214, 157, 133) }, // 棕色
                    { "number", Color.FromArgb(181, 206, 168) }, // 蓝绿色
                    { "property", Color.FromArgb(156, 220, 254) }, // 浅蓝色
                    { "selector", Color.FromArgb(215, 186, 125) }, // 金色
                };
            }
            else
            {
                themeColors = new Dictionary<string, Color>
                {
                    { "keyword", Color.FromArgb(0, 0, 255) }, // 蓝色
                    { "comment", Color.FromArgb(0, 128, 0) }, // 绿色
                    { "string", Color.FromArgb(163, 21, 21) }, // 棕色
                    { "number", Color.FromArgb(9, 134, 88) }, // 蓝绿色
                    { "property", Color.FromArgb(0, 134, 209) }, // 浅蓝色
                    { "selector", Color.FromArgb(168, 99, 0) }, // 褐色
                };
            }
        }

        private void InitializePatterns()
        {
            languagePatterns = new Dictionary<
                string,
                List<(string pattern, RegexOptions options, string type)>
            >
            {
                {
                    "csharp",
                    new List<(string, RegexOptions, string)>
                    {
                        // 注释 - 放在最前面优先匹配
                        (@"/\*[\s\S]*?\*/", RegexOptions.None, "comment"),
                        (@"//[^\n]*", RegexOptions.None, "comment"),
                        // 字符串
                        (@"@""([^""]|"""")*""|""([^""\n\\]|\\.)*""", RegexOptions.None, "string"),
                        // 数字
                        (@"\b\d*\.?\d+([eE][-+]?\d+)?\b", RegexOptions.None, "number"),
                        // 关键字
                        (
                            @"\b(abstract|as|base|bool|break|byte|case|catch|char|checked|class|const|continue|decimal|default|delegate|do|double|else|enum|event|explicit|extern|false|finally|fixed|float|for|foreach|goto|if|implicit|in|int|interface|internal|is|lock|long|namespace|new|null|object|operator|out|override|params|private|protected|public|readonly|ref|return|sbyte|sealed|short|sizeof|stackalloc|static|string|struct|switch|this|throw|true|try|typeof|uint|ulong|unchecked|unsafe|ushort|using|virtual|void|volatile|while)\b",
                            RegexOptions.None,
                            "keyword"
                        ),
                    }
                },
                {
                    "python",
                    new List<(string, RegexOptions, string)>
                    {
                        // 注释 - 放在最前面优先匹配
                        (@"#.*?$", RegexOptions.Multiline, "comment"),
                        // 字符串
                        (
                            @"(?:[ruf]|rf|fr)?(?:'''[\s\S]*?'''|\""""""[\s\S]*?\""""""|\x27[^\x27\n\\]*(?:\\.[^\x27\n\\]*)*\x27|\""[^\""\n\\]*(?:\\.[^\""\n\\]*)*\"")",
                            RegexOptions.None,
                            "string"
                        ),
                        // 数字
                        (@"\b\d*\.?\d+([eE][-+]?\d+)?\b", RegexOptions.None, "number"),
                        (
                            @"\b(and|as|assert|async|await|break|class|continue|def|del|elif|else|except|False|finally|for|from|global|if|import|in|is|lambda|None|nonlocal|not|or|pass|raise|return|True|try|while|with|yield)\b",
                            RegexOptions.None,
                            "keyword"
                        ),
                    }
                },
                {
                    "javascript",
                    new List<(string, RegexOptions, string)>
                    {
                        // 注释 - 放在最前面优先匹配
                        (@"/\*[\s\S]*?\*/|//.*?$", RegexOptions.Multiline, "comment"),
                        // 字符串
                        (
                            @"`(?:[^`\\]|\\.)*`|'(?:[^'\\]|\\.)*'|""(?:[^""\\]|\\.)*""",
                            RegexOptions.None,
                            "string"
                        ),
                        // 数字
                        (@"\b\d*\.?\d+([eE][-+]?\d+)?\b", RegexOptions.None, "number"),
                        (
                            @"\b(async|await|break|case|catch|class|const|continue|debugger|default|delete|do|else|export|extends|finally|for|function|if|import|in|instanceof|new|return|super|switch|this|throw|try|typeof|var|void|while|with|yield|let)\b",
                            RegexOptions.None,
                            "keyword"
                        ),
                    }
                },
                {
                    "matlab",
                    new List<(string, RegexOptions, string)>
                    {
                        // 注释 - 放在最前面优先匹配
                        (@"%.*?$", RegexOptions.Multiline, "comment"),
                        // 字符串 (支持单引号和双引号)
                        (@"(?:""[^""\n]*""|'[^'\n]*')", RegexOptions.None, "string"),
                        // 数字
                        (@"\b\d*\.?\d+([eE][-+]?\d+)?\b", RegexOptions.None, "number"),
                        // 关键字
                        (
                            @"\b(break|case|catch|classdef|continue|else|elseif|end|for|function|global|if|otherwise|parfor|persistent|return|switch|try|while|clear|close|load|save|figure|plot|xlabel|ylabel|title|grid|hold|zeros|ones|rand|eye|disp|input|fprintf|strcmp|length|size|max|min|sum|mean|std|find|sort|reshape)\b",
                            RegexOptions.None,
                            "keyword"
                        ),
                    }
                },
                {
                    "css",
                    new List<(string, RegexOptions, string)>
                    {
                        // 注释 - 放在最前面优先匹配
                        (@"/\*[\s\S]*?\*/", RegexOptions.None, "comment"),
                        // CSS选择器 - 在大括号前的任何非大括号内容
                        (@"[^{}]+(?=\s*{)", RegexOptions.None, "selector"),
                        // CSS属性
                        (@"[\w-]+(?=\s*:)", RegexOptions.None, "property"),
                        // 字符串
                        (@"'[^']*'|""[^""]*""", RegexOptions.None, "string"),
                        // 数字和单位
                        (
                            @"\b\d*\.?\d+(?:px|em|rem|vh|vw|%|s|ms|deg|rad|turn)?\b",
                            RegexOptions.None,
                            "number"
                        ),
                        // 关键字
                        (
                            @"\b(important|inherit|initial|unset|none|auto|hidden|visible|block|inline|flex|grid|absolute|relative|fixed|static|left|right|top|bottom|center|justify|stretch|wrap|nowrap|solid|dashed|dotted)\b",
                            RegexOptions.None,
                            "keyword"
                        ),
                    }
                },
                {
                    "html",
                    new List<(string, RegexOptions, string)>
                    {
                        // 注释 - 放在最前面优先匹配
                        (@"<!--[\s\S]*?-->", RegexOptions.None, "comment"),
                        // 字符串
                        (@"'[^']*'|""[^""]*""", RegexOptions.None, "string"),
                        // HTML标签
                        (
                            @"</?\w+(?:\s+\w+(?:\s*=\s*(?:"".*?""|'.*?'|[\^'""\s]\S+))?)*\s*/?>",
                            RegexOptions.None,
                            "keyword"
                        ),
                        // 数字
                        (@"\b\d+\b", RegexOptions.None, "number"),
                    }
                },
                {
                    "r",
                    new List<(string, RegexOptions, string)>
                    {
                        // 注释 - 放在最前面优先匹配
                        (@"#.*?$", RegexOptions.Multiline, "comment"),
                        // 字符串 (支持单引号和双引号)
                        (
                            @"(?:""[^""\n\\]*(?:\\.[^""\n\\]*)*""|'[^'\n\\]*(?:\\.[^'\n\\]*)*')",
                            RegexOptions.None,
                            "string"
                        ),
                        // 数字
                        (@"\b\d*\.?\d+([eE][-+]?\d+)?L?\b", RegexOptions.None, "number"),
                        // 关键字和函数
                        (
                            @"\b(if|else|for|in|while|function|repeat|next|break|TRUE|FALSE|NULL|Inf|NaN|NA|NA_integer_|NA_real_|NA_complex_|NA_character_|c|list|data\.frame|matrix|array|factor|length|names|dim|class|str|summary|head|tail|print|cat|paste|paste0|substr|nchar|grep|gsub|which|is\.na|is\.null|is\.numeric|is\.character|is\.logical|as\.numeric|as\.character|as\.logical|mean|median|sd|var|min|max|sum|apply|lapply|sapply|mapply|tapply|aggregate|merge|rbind|cbind|subset|sort|order|unique|duplicated|table|plot|hist|boxplot|barplot|pie|lines|points|abline|legend|title|xlabel|ylabel|par|dev\.new|dev\.off|png|pdf|jpeg|library|require|install\.packages|source|load|save|write\.csv|read\.csv|read\.table|write\.table)\b",
                            RegexOptions.None,
                            "keyword"
                        ),
                    }
                },
            };
        }

        private void InitializeAliases()
        {
            languageAliases = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "js", "javascript" },
                { "c#", "csharp" },
                { "cs", "csharp" },
                { "py", "python" },
                { "m", "matlab" },
                { "css", "css" },
                { "htm", "html" },
                { "R", "r" },
            };
        }

        private int GetActualPosition(string text, int position)
        {
            // 计算到指定位置前的换行符数量
            int newlineCount = text.Substring(0, position).Count(c => c == '\n' || c == '\r');
            // PowerPoint中每个换行符只算一个字符，而在字符串中\r\n算两个字符
            int adjustment = text.Substring(0, position).Count(c => c == '\r');
            return position - adjustment;
        }

        public void ApplyHighlighting(PowerPoint.Shape textBox, string code, string language)
        {
            // Normalize language name
            string normalizedLanguage = language.ToLower();
            if (languageAliases.ContainsKey(normalizedLanguage))
            {
                normalizedLanguage = languageAliases[normalizedLanguage];
            }

            if (!languagePatterns.ContainsKey(normalizedLanguage))
                return;

            processedRanges.Clear();
            var patterns = languagePatterns[normalizedLanguage];

            // 收集所有匹配项
            var allMatches = new List<(Match match, string type, int priority)>();

            foreach (var (pattern, options, type) in patterns)
            {
                var regex = new Regex(pattern, options);
                var matches = regex.Matches(code);

                // 设置优先级：注释最高优先级
                int priority = type == "comment" ? 0 : 1;

                foreach (Match match in matches)
                {
                    allMatches.Add((match, type, priority));
                }
            }

            // 按优先级排序，优先级数字越小越优先
            allMatches.Sort((a, b) => a.priority.CompareTo(b.priority));

            // 记录已处理的字符范围
            var processedIntervals = new List<(int start, int end)>();

            foreach (var (match, type, priority) in allMatches)
            {
                try
                {
                    int matchStart = match.Index;
                    int matchEnd = match.Index + match.Length;

                    // 检查是否与已处理的区域重叠
                    bool hasOverlap = false;
                    foreach (var (start, end) in processedIntervals)
                    {
                        if (matchStart < end && matchEnd > start)
                        {
                            hasOverlap = true;
                            break;
                        }
                    }

                    if (hasOverlap)
                        continue;

                    // 计算实际的开始位置和长度
                    int actualStart = GetActualPosition(code, matchStart);
                    int actualLength = GetActualPosition(code, matchEnd) - actualStart;

                    var range = textBox.TextFrame.TextRange.Characters(
                        actualStart + 1,
                        actualLength
                    );
                    range.Font.Color.RGB = ColorTranslator.ToOle(themeColors[type]);

                    // 记录已处理的区间
                    processedIntervals.Add((matchStart, matchEnd));
                }
                catch (Exception)
                {
                    continue;
                }
            }
        }
    }
}
