using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class ParameterCanBeByValInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public ParameterCanBeByValInspectionResult(IInspection inspection, RubberduckParserState state, Declaration target, ParserRuleContext context, QualifiedMemberName qualifiedName)
            : base(inspection, qualifiedName.QualifiedModuleName, context, target)
        {
            _quickFixes = new CodeInspectionQuickFix[]
            {
                new PassParameterByValueQuickFix(state, Target, Context, QualifiedSelection),
                new IgnoreOnceQuickFix(Context, QualifiedSelection, inspection.AnnotationName)
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get { return string.Format(InspectionsUI.ParameterCanBeByValInspectionResultFormat, Target.IdentifierName); }
        }
    }
}
