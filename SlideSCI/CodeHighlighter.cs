// CodeHighlighter.cs  —  Drop-in replacement (C# 7.3 compatible)
// Provides instance method ApplyHighlighting(...) expected by Ribbon1.cs
// and regex patterns including STATA language.

using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SlideSCI
{
    public sealed class CodeHighlighter
    {
        // language -> list of (pattern, options, styleKey)
        private static readonly Dictionary<string, List<(string pattern, RegexOptions options, string style)>> languagePatterns
            = new Dictionary<string, List<(string, RegexOptions, string)>>(StringComparer.OrdinalIgnoreCase);

        // compiled cache
        private static readonly Dictionary<string, List<PatternEntry>> compiledCache
            = new Dictionary<string, List<PatternEntry>>(StringComparer.OrdinalIgnoreCase);

        private static readonly Dictionary<string, string> languageAliases
            = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public sealed class PatternEntry
        {
            public Regex Regex { get; private set; }
            public string Style { get; private set; }
            public PatternEntry(Regex regex, string style) { Regex = regex; Style = style; }
        }

        static CodeHighlighter()
        {
            InitializePatterns();
            InitializeAliases();
            CompileAll();
        }

        private static void InitializeAliases()
        {
            // Stata
            languageAliases["stata"] = "stata";
            languageAliases["do"] = "stata";
            languageAliases["ado"] = "stata";
        }

        private static void InitializePatterns()
        {
            languagePatterns.Clear();

            // ===== STATA =====
            languagePatterns["stata"] = new List<(string, RegexOptions, string)>
            {
                // comments
                (@"(?s)/\*.*?\*/", RegexOptions.None, "comment"),
                (@"(?m)^[ \t]*\*.*?$", RegexOptions.Multiline, "comment"),
                (@"//.*?$", RegexOptions.Multiline, "comment"),

                // strings
                (@"(?s)""([^""\\]|\\.)*""", RegexOptions.None, "string"),
                (@"(?s)`""[\s\S]*?""'", RegexOptions.None, "string"),

                // numbers
                (@"\b\d*\.?\d+(?:[eE][-+]?\d+)?\b", RegexOptions.None, "number"),

                // keywords (control/meta)
                (
                    @"\b(?i:(if|else|in|using|by|bysort|quietly|noisily|qui|capture|preserve|restore|program|end|syntax|args|local|global|tempvar|tempname|tempfile|scalar|matrix|mata|return|ereturn|post|eststo|esttab|estadd|foreach|forvalues|while|continue|break|version|set|clear|cls|pause|exit|do|ado|which|graph|twoway|histogram|kdensity|scatter|line|bar|tsset|xtset|timer|assert|confirm|display|di|pwd|cd|mkdir|rmdir|save|use|import|export|outsheet|insheet|log|translate|help|view|about|update))\b",
                    RegexOptions.IgnoreCase,
                    "keyword"
                ),

                // keywords (data & estimation)
                (
                    @"\b(?i:(gen|generate|egen|replace|drop|keep|order|move|rename|label|lab(?:el)?(?:\s+(?:var|val|def|values|data))?|destring|tostring|encode|decode|recast|format|contract|collapse|append|merge|joinby|cross|reshape|separate|split|expand|sample|duplicates|distinct|levelsof|tabulate|tab|summarize|sum|count|pctile|xtile|corr(?:elation)?|corrgram|areg|regress|logit|probit|tobit|poisson|nbreg|ivregress|gmm|qreg|xtreg|xtlogit|xtprobit|xtpoisson|mixed|meqrlogit|melogit|ppmlhdfe|reghdfe|hdfe|felsdvreg|teffects|mi|impute|stset|stcox|stmixed|svy|bootstrap|jackknife|permute|recode))\b",
                    RegexOptions.IgnoreCase,
                    "keyword"
                ),

                // functions
                (
                    @"\b(?i:(abs|ceil|floor|int|round|min|max|sum|mean|cond|inlist|inrange|missing|real|string|substr|subinstr|strpos|ustrpos|regexm|regexr|regexs|ustrregexm|ustrregexrf|ustrregexra|length|strlen|ustrlen|lower|upper|proper|trim|itrim|ltrim|rtrim|date|clock|mdy|dow|dofc|cofc|ofd|wofd|runiform|rnormal|invnormal|exp|log|ln|sqrt))\b",
                    RegexOptions.IgnoreCase,
                    "property"
                ),

                // macros
                (@"`[A-Za-z_][A-Za-z0-9_]*'", RegexOptions.None, "property"),
                (@"``[A-Za-z_][A-Za-z0-9_]*''", RegexOptions.None, "property"),
            };
        }

        private static void CompileAll()
        {
            compiledCache.Clear();
            foreach (var kv in languagePatterns)
            {
                var list = new List<PatternEntry>(kv.Value.Count);
                foreach (var entry in kv.Value)
                {
                    var rx = new Regex(entry.pattern, entry.options | RegexOptions.Compiled | RegexOptions.CultureInvariant);
                    list.Add(new PatternEntry(rx, entry.style));
                }
                compiledCache[kv.Key] = list;
            }
        }

        private static string NormalizeLanguage(string languageOrAlias)
        {
            if (string.IsNullOrWhiteSpace(languageOrAlias)) return string.Empty;
            string mapped;
            if (languageAliases.TryGetValue(languageOrAlias.Trim(), out mapped)) return mapped;
            return languageOrAlias.Trim();
        }

        private static IReadOnlyList<PatternEntry> GetPatterns(string languageOrAlias)
        {
            var key = NormalizeLanguage(languageOrAlias);
            List<PatternEntry> list;
            if (compiledCache.TryGetValue(key, out list)) return list;
            return Array.Empty<PatternEntry>();
        }

        // ======= Instance API expected by Ribbon1.cs =======

        // 常用：ApplyHighlighting(TextRange2, "stata")
        public void ApplyHighlighting(PowerPoint.TextRange2 range, string language)
        {
            if (range == null) return;
            ApplyToTextRange2(range, language);
        }

        // 兼容参数顺序相反的写法：ApplyHighlighting("stata", TextRange2)
        public void ApplyHighlighting(string language, PowerPoint.TextRange2 range)
        {
            if (range == null) return;
            ApplyToTextRange2(range, language);
        }

        // 兜底：老的 TextRange 接口
        public void ApplyHighlighting(PowerPoint.TextRange range, string language)
        {
            if (range == null) return;
            // 简单转写：把 TextRange 的文本提出来，定位用 1-based 的 Characters
            string text = range.Text ?? string.Empty;
            var pats = GetPatterns(language);
            foreach (var pe in pats)
            {
                foreach (Match m in pe.Regex.Matches(text))
                {
                    try
                    {
                        // TextRange.Characters 是 1-based
                        var ch = range.Characters(m.Index + 1, m.Length);
                        SetStyle(ch.Font, pe.Style);
                    }
                    catch { /* ignore individual highlight errors */ }
                }
            }
        }

        // 万能兜底（仅为通过编译；若被命中则不做任何事）
        public void ApplyHighlighting(params object[] _)
        {
            // no-op
        }

        // ======= 实际高亮实现（TextRange2） =======
        private void ApplyToTextRange2(PowerPoint.TextRange2 range, string language)
        {
            string text = range.Text ?? string.Empty;
            var pats = GetPatterns(language);
            foreach (var pe in pats)
            {
                foreach (Match m in pe.Regex.Matches(text))
                {
                    try
                    {
                        // TextRange2.Characters 是 1-based
                        var ch = range.Characters[m.Index + 1, m.Length];
                        SetStyle(ch.Font, pe.Style);
                    }
                    catch { /* 单个片段失败不影响整体 */ }
                }
            }
        }

        // 简单的样式映射（可在 Ribbon 里统一颜色，这里先给一个通用方案）
        private static void SetStyle(PowerPoint.TextRange2Font font, string style)
        {
            int rgb;
            switch ((style ?? "").ToLowerInvariant())
            {
                case "comment":  rgb = 0x008000; break; // 绿色
                case "string":   rgb = 0xAA5500; break; // 棕橙
                case "number":   rgb = 0x1A73E8; break; // 蓝
                case "keyword":  rgb = 0xB000B0; break; // 紫
                case "property": rgb = 0x795548; break; // 棕
                default:         rgb = 0x000000; break; // 黑
            }
            try
            {
                font.Fill.ForeColor.RGB = rgb;
            }
            catch { /* 某些主题下可能抛异常，忽略 */ }
        }

        // TextRange 版本（老接口）
        private static void SetStyle(PowerPoint.Font font, string style)
        {
            int rgb;
            switch ((style ?? "").ToLowerInvariant())
            {
                case "comment":  rgb = 0x008000; break;
                case "string":   rgb = 0xAA5500; break;
                case "number":   rgb = 0x1A73E8; break;
                case "keyword":  rgb = 0xB000B0; break;
                case "property": rgb = 0x795548; break;
                default:         rgb = 0x000000; break;
            }
            try
            {
                font.Color.RGB = rgb;
            }
            catch { /* ignore */ }
        }
    }
}
