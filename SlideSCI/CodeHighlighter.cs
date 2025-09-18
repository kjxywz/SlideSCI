// CodeHighlighter.cs
// Drop-in: provides regex-based syntax highlighting patterns.
// NOTE: Adjust namespace/class name to match your project if needed.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace SlideSCI.Code
{
    /// <summary>
    /// Central registry for code syntax patterns.
    /// Provides compiled Regex + style keys for a given language (with simple aliasing).
    /// Styles follow keys: "keyword", "comment", "string", "number", "property".
    /// </summary>
    public static class CodeHighlighter
    {
        // Raw patterns: language -> list of (regex pattern, options, styleKey)
        private static readonly Dictionary<string, List<(string pattern, RegexOptions options, string style)>> languagePatterns
            = new(StringComparer.OrdinalIgnoreCase);

        // Compiled cache: languageKey -> list of compiled entries
        private static readonly Dictionary<string, List<PatternEntry>> compiledCache
            = new(StringComparer.OrdinalIgnoreCase);

        // Aliases: like "do"/"ado" -> "stata"
        private static readonly Dictionary<string, string> languageAliases
            = new(StringComparer.OrdinalIgnoreCase);

        static CodeHighlighter()
        {
            InitializePatterns();
            InitializeAliases();
            CompileAll();
        }

        /// <summary>
        /// The compiled regex + style entry.
        /// </summary>
        public sealed class PatternEntry
        {
            public Regex Regex { get; }
            public string Style { get; }

            public PatternEntry(Regex regex, string style)
            {
                Regex = regex;
                Style = style;
            }
        }

        /// <summary>
        /// Return compiled patterns for a language (resolves aliases).
        /// Returns empty list if not found.
        /// </summary>
        public static IReadOnlyList<PatternEntry> GetPatterns(string languageOrAlias)
        {
            if (string.IsNullOrWhiteSpace(languageOrAlias))
                return Array.Empty<PatternEntry>();

            var key = GetLanguageKey(languageOrAlias);
            if (compiledCache.TryGetValue(key, out var list))
                return list;

            return Array.Empty<PatternEntry>();
        }

        /// <summary>
        /// True if the language (or alias) is supported.
        /// </summary>
        public static bool HasLanguage(string languageOrAlias)
        {
            var key = GetLanguageKey(languageOrAlias);
            return compiledCache.ContainsKey(key);
        }

        /// <summary>
        /// Normalize language by alias map.
        /// </summary>
        public static string GetLanguageKey(string languageOrAlias)
        {
            if (string.IsNullOrWhiteSpace(languageOrAlias))
                return languageOrAlias ?? string.Empty;

            if (languageAliases.TryGetValue(languageOrAlias.Trim(), out var mapped))
                return mapped;

            return languageOrAlias.Trim();
        }

        private static void CompileAll()
        {
            compiledCache.Clear();
            foreach (var kv in languagePatterns)
            {
                var list = new List<PatternEntry>(kv.Value.Count);
                foreach (var (pattern, opts, style) in kv.Value)
                {
                    // Compiled + CultureInvariant is fine for code
                    var regex = new Regex(pattern, opts | RegexOptions.Compiled | RegexOptions.CultureInvariant);
                    list.Add(new PatternEntry(regex, style));
                }
                compiledCache[kv.Key] = list;
            }
        }

        private static void InitializeAliases()
        {
            // --- Stata ---
            languageAliases["stata"] = "stata";
            languageAliases["do"] = "stata";
            languageAliases["ado"] = "stata";

            // (Optional) Keep your existing aliases here if needed for other languages.
            // e.g., languageAliases["py"] = "python";
        }

        private static void InitializePatterns()
        {
            languagePatterns.Clear();

            // ========== STATA ==========
            // Ordering matters: comment > string > number > keyword > function/property > macros
            languagePatterns["stata"] = new List<(string, RegexOptions, string)>
            {
                // --- Comments ---
                // Block comment: /* ... */
                (@"(?s)/\*.*?\*/", RegexOptions.None, "comment"),
                // Line comment: leading * (common Stata style)
                (@"(?m)^[ \t]*\*.*?$", RegexOptions.Multiline, "comment"),
                // Line-end comment: // to end-of-line
                (@"//.*?$", RegexOptions.Multiline, "comment"),

                // --- Strings ---
                // Normal double-quoted string (allow escaped chars loosely)
                (@"(?s)""([^""\\]|\\.)*""", RegexOptions.None, "string"),
                // Compound double quotes: `"... "'  (local macros often use these)
                (@"(?s)`""[\s\S]*?""'", RegexOptions.None, "string"),

                // --- Numbers (integer/float/scientific) ---
                (@"\b\d*\.?\d+(?:[eE][-+]?\d+)?\b", RegexOptions.None, "number"),

                // --- Keywords: control/meta commands ---
                (
                    @"\b(?i:(if|else|in|using|by|bysort|quietly|noisily|qui|capture|preserve|restore|program|end|syntax|args|local|global|tempvar|tempname|tempfile|scalar|matrix|mata|return|ereturn|post|eststo|esttab|estadd|foreach|forvalues|while|continue|break|version|set|clear|cls|pause|exit|do|ado|which|graph|twoway|histogram|kdensity|scatter|line|bar|tsset|xtset|timer|assert|confirm|display|di|pwd|cd|mkdir|rmdir|save|use|import|export|outsheet|insheet|log|translate|help|view|about|update))\b",
                    RegexOptions.IgnoreCase,
                    "keyword"
                ),

                // --- Keywords: data & estimation commands (common set; editable/extendable) ---
                (
                    @"\b(?i:(gen|generate|egen|replace|drop|keep|order|move|rename|label|lab(?:el)?(?:\s+(?:var|val|def|values|data))?|destring|tostring|encode|decode|recast|format|contract|collapse|append|merge|joinby|cross|reshape|separate|split|expand|sample|duplicates|distinct|levelsof|tabulate|tab|summarize|sum|count|pctile|xtile|corr(?:elation)?|corrgram|areg|regress|logit|probit|tobit|poisson|nbreg|ivregress|gmm|qreg|xtreg|xtlogit|xtprobit|xtpoisson|mixed|meqrlogit|melogit|ppmlhdfe|reghdfe|hdfe|felsdvreg|teffects|mi|impute|stset|stcox|stmixed|svy|bootstrap|jackknife|permute|recode))\b",
                    RegexOptions.IgnoreCase,
                    "keyword"
                ),

                // --- Built-in functions (assign "property" style to match many themes) ---
                (
                    @"\b(?i:(abs|ceil|floor|int|round|min|max|sum|mean|cond|inlist|inrange|missing|real|string|substr|subinstr|strpos|ustrpos|regexm|regexr|regexs|ustrregexm|ustrregexrf|ustrregexra|length|strlen|ustrlen|lower|upper|proper|trim|itrim|ltrim|rtrim|date|clock|mdy|dow|dofc|cofc|ofd|wofd|runiform|rnormal|invnormal|exp|log|ln|sqrt))\b",
                    RegexOptions.IgnoreCase,
                    "property"
                ),

                // --- Macro references (local/global names in backticks) ---
                // `name'
                (@"`[A-Za-z_][A-Za-z0-9_]*'", RegexOptions.None, "property"),
                // ``name'' (double-backtick / double-quote style)
                (@"``[A-Za-z_][A-Za-z0-9_]*''", RegexOptions.None, "property"),
            };

            // === (Optional) Keep/restore other languages here if your project needs them. ===
            // e.g., languagePatterns["python"] = ...
            // This file focuses on adding Stata; other entries can remain as in your original.
        }
    }
}
