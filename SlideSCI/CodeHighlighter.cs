// CodeHighlighter.cs — TextRange-only build (C# 7.3 compatible)
// No direct reference to TextRange2/TextRange2Font: works on CI with Office15 PIA.
// Provides instance ApplyHighlighting(...) expected by Ribbon1.cs.

using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text.RegularExpressions;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SlideSCI
{
    public sealed class CodeHighlighter
    {
        // language -> list of (pattern, options, styleKey)
        private static readonly Dictionary<string, List<(string pattern, RegexOptions options, string style)>> languagePatterns
            = new Dictionary<string, List<(string, RegexOptions, string)>>(StringComparer.OrdinalIgnoreCase);

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
                    RegexOptions.IgnoreCase, "keyword"
                ),
                // keywords (data & estimation)
                (
                    @"\b(?i:(gen|generate|egen|replace|drop|keep|order|move|rename|label|lab(?:el)?(?:\s+(?:var|val|def|values|data))?|destring|tostring|encode|decode|recast|format|contract|collapse|append|merge|joinby|cross|reshape|separate|split|expand|sample|duplicates|distinct|levelsof|tabulate|tab|summarize|sum|count|pctile|xtile|corr(?:elation)?|corrgram|areg|regress|logit|probit|tobit|poisson|nbreg|ivregress|gmm|qreg|xtreg|xtlogit|xtprobit|xtpoisson|mixed|meqrlogit|melogit|ppmlhdfe|reghdfe|hdfe|felsdvreg|teffects|mi|impute|stset|stcox|stmixed|svy|bootstrap|jackknife|permute|recode))\b",
                    RegexOptions.IgnoreCase, "keyword"
                ),
                // functions -> property
                (
                    @"\b(?i:(abs|ceil|floor|int|round|min|max|sum|mean|cond|inlist|inrange|missing|real|string|substr|subinstr|strpos|ustrpos|regexm|regexr|regexs|ustrregexm|ustrregexrf|ustrregexra|length|strlen|ustrlen|lower|upper|proper|trim|itrim|ltrim|rtrim|date|clock|mdy|dow|dofc|cofc|ofd|wofd|runiform|rnormal|invnormal|exp|log|ln|sqrt))\b",
                    RegexOptions.IgnoreCase, "property"
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

        private static IReadOnlyList<PatternEntry> GetPatterns(string languageOrAlias)
        {
            if (string.IsNullOrWhiteSpace(languageOrAlias)) return Array.Empty<PatternEntry>();
            string mapped;
            if (languageAliases.TryGetValue(languageOrAlias.Trim(), out mapped)) languageOrAlias = mapped;
            List<PatternEntry> list;
            if (compiledCache.TryGetValue(languageOrAlias.Trim(), out list)) return list;
            return Array.Empty<PatternEntry>();
        }

        // ========= Public API expected by Ribbon1 =========

        // 1) 直接支持 TextRange
        public void ApplyHighlighting(PowerPoint.TextRange range, string language)
        {
            if (range == null) return;
            ApplyToTextRange(range, language);
        }

        // 2) 兜底：任何类型（例如 TextRange2）都会落到这里
        public void ApplyHighlighting(object range, string language)
        {
            if (range == null) return;

            // 优先处理 TextRange
            var tr = range as PowerPoint.TextRange;
            if (tr != null) { ApplyToTextRange(tr, language); return; }

            // 兼容 TextRange2（CI 上没有类型定义，所以用反射）
            try
            {
                var type = range.GetType();                         // e.g. TextRange2
                var textProp = type.GetProperty("Text");
                var charsIndexer = type.GetProperty("Characters");  // indexer-like property
                if (textProp == null || charsIndexer == null) return;

                string text = textProp.GetValue(range, null) as string ?? string.Empty;
                var pats = GetPatterns(language);
                foreach (var pe in pats)
                {
                    foreach (Match m in pe.Regex.Matches(text))
                    {
                        try
                        {
                            // TextRange2.Characters is 1-based, length as second arg
                            var ch = charsIndexer.GetValue(range, new object[] { m.Index + 1, m.Length });

                            // ch.Font 可能是 TextRange2Font，继续用反射设置 Fill.ForeColor.RGB
                            var fontProp = ch.GetType().GetProperty("Font");
                            if (fontProp != null)
                            {
                                var fontObj = fontProp.GetValue(ch, null);
                                // font.Fill.ForeColor.RGB
                                var fillProp = fontObj.GetType().GetProperty("Fill");
                                if (fillProp != null)
                                {
                                    var fill = fillProp.GetValue(fontObj, null);
                                    var fcProp = fill.GetType().GetProperty("ForeColor");
                                    if (fcProp != null)
                                    {
                                        var fc = fcProp.GetValue(fill, null);
                                        var rgbProp = fc.GetType().GetProperty("RGB");
                                        if (rgbProp != null) rgbProp.SetValue(fc, StyleRgb(pe.Style), null);
                                    }
                                }
                            }
                        }
                        catch { /* ignore single failure */ }
                    }
                }
            }
            catch { /* ignore */ }
        }

        // ========== Internal implementations ==========

        private void ApplyToTextRange(PowerPoint.TextRange range, string language)
        {
            string text = range.Text ?? string.Empty;
            var pats = GetPatterns(language);
            foreach (var pe in pats)
            {
                foreach (Match m in pe.Regex.Matches(text))
                {
                    try
                    {
                        // TextRange.Characters is 1-based
                        var ch = range.Characters(m.Index + 1, m.Length);
                        // TextRange.Font.Color.RGB
                        ch.Font.Color.RGB = StyleRgb(pe.Style);
                    }
                    catch { /* ignore */ }
                }
            }
        }

        private static int StyleRgb(string style)
        {
            switch ((style ?? "").ToLowerInvariant())
            {
                case "comment":  return 0x008000; // green
                case "string":   return 0xAA5500; // brown-orange
                case "number":   return 0x1A73E8; // blue
                case "keyword":  return 0xB000B0; // purple
                case "property": return 0x795548; // brown
                default:         return 0x000000; // black
            }
        }
    }
}
