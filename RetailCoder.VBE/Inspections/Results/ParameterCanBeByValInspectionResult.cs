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
        private IEnumerable<QuickFixBase> _quickFixes;
        private readonly RubberduckParserState _state;

        public ParameterCanBeByValInspectionResult(IInspection inspection, RubberduckParserState state, Declaration target, ParserRuleContext context, QualifiedMemberName qualifiedName)
            : base(inspection, qualifiedName.QualifiedModuleName, context, target)
        {
            _state = state;
        }

        public override IEnumerable<QuickFixBase> QuickFixes
        {
            get
            {
                return _quickFixes ?? (_quickFixes = new QuickFixBase[]
                {
                    new PassParameterByValueQuickFix(_state, Target, Context, QualifiedSelection),
                    new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName)
                });
            }
        }

        public override string Description
        {
            get { return string.Format(InspectionsUI.ParameterCanBeByValInspectionResultFormat, Target.IdentifierName).Captialize(); }
        }
    }
}
