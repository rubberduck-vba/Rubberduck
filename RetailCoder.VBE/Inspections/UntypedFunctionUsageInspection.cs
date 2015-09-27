using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Nodes;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class UntypedFunctionUsageInspection : IInspection
    {
        public UntypedFunctionUsageInspection()
        {
            Severity = CodeInspectionSeverity.Hint;
        }

        public string Name { get { return "UntypedFunctionUsageInspection"; } }
        public string Description { get { return RubberduckUI.UntypedFunctionUsage_; } }
        public CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }
        public CodeInspectionSeverity Severity { get; set; }

        private readonly string[] _tokens = new string[]
        {
            Tokens.Error,
            Tokens.Hex,
            Tokens.Oct,
            Tokens.Str,
            Tokens.CurDir,
            Tokens.Command,
            Tokens.Environ,
            Tokens.Chr,
            Tokens.ChrW,
            Tokens.Format,
            Tokens.LCase,
            Tokens.Left,
            Tokens.LeftB,
            Tokens.LTrim,
            Tokens.Mid,
            Tokens.MidB,
            Tokens.Trim,
            Tokens.Right,
            Tokens.RightB,
            Tokens.RTrim,
            Tokens.UCase
        };

        public IEnumerable<CodeInspectionResultBase> GetInspectionResults(VBProjectParseResult parseResult)
        {
            var declarations = parseResult.Declarations.Items
                .Where(item => item.IsBuiltIn && item.Accessibility == Accessibility.Global && _tokens.Contains(item.IdentifierName));

            return declarations.SelectMany(declaration => declaration.References
                .Select(item => new UntypedFunctionUsageInspectionResult(this, string.Format(Description, declaration.IdentifierName), item.QualifiedModuleName, item.Context)));
        }
    }

    public class UntypedFunctionUsageInspectionResult : CodeInspectionResultBase
    {
        public UntypedFunctionUsageInspectionResult(IInspection inspection, string result, QualifiedModuleName qualifiedName, ParserRuleContext context) 
            : base(inspection, result, qualifiedName, context)
        {
        }

        //todo: override QuickFixes getter
        //public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get; private set; }
    }
}