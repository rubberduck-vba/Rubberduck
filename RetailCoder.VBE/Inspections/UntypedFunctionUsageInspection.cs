using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class UntypedFunctionUsageInspection : InspectionBase
    {
        public UntypedFunctionUsageInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Hint)
        {
        }

        public override string Meta { get { return InspectionsUI.UntypedFunctionUsageInspectionMeta; } }
        public override string Description { get { return InspectionsUI.UntypedFunctionUsageInspectionResultFormat; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }

        private readonly string[] _tokens = {
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

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            var declarations = BuiltInDeclarations
                // note: these *should* be functions, but somehow they're not defined as such
                .Where(item =>
                        _tokens.Any(token => item.IdentifierName == token || item.IdentifierName == "_B_var_" + token) &&
                        item.References.Any(reference => _tokens.Contains(reference.IdentifierName)));

            return declarations.SelectMany(declaration => declaration.References
                .Where(item => _tokens.Contains(item.IdentifierName))
                .Select(item => new UntypedFunctionUsageInspectionResult(this, item)));
        }
    }
}
