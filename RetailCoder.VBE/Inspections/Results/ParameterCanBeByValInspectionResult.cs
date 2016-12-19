using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Results
{
    public class ParameterCanBeByValInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<QuickFixBase> _quickFixes;

        public ParameterCanBeByValInspectionResult(IInspection inspection, RubberduckParserState state, Declaration target, ParserRuleContext context, QualifiedMemberName qualifiedName)
            : base(inspection, qualifiedName.QualifiedModuleName, context, target)
        {
            _quickFixes = new QuickFixBase[]
            {
                new PassParameterByValueQuickFix(state, Target, Context, QualifiedSelection),
                new IgnoreOnceQuickFix(Context, QualifiedSelection, inspection.AnnotationName)
            };
        }

        public override IEnumerable<QuickFixBase> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get { return string.Format(InspectionsUI.ParameterCanBeByValInspectionResultFormat, Target.IdentifierName).Captialize(); }
        }
    }
}
