using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class ParameterNotUsedInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public ParameterNotUsedInspectionResult(IInspection inspection, Declaration target,
            ParserRuleContext context, QualifiedMemberName qualifiedName, bool isInterfaceImplementation, 
            RemoveParametersRefactoring refactoring, RubberduckParserState state)
            : base(inspection, qualifiedName.QualifiedModuleName, context, target)
        {
            _quickFixes = isInterfaceImplementation ? new CodeInspectionQuickFix[] {} : new CodeInspectionQuickFix[]
            {
                new RemoveUnusedParameterQuickFix(Context, QualifiedSelection, refactoring, state),
                new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName), 
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get { return string.Format(InspectionsUI.ParameterNotUsedInspectionResultFormat, Target.IdentifierName); }
        }
    }

    public class RemoveUnusedParameterQuickFix : CodeInspectionQuickFix
    {
        private readonly RemoveParametersRefactoring _quickFixRefactoring;
        private readonly RubberduckParserState _state;

        public RemoveUnusedParameterQuickFix(ParserRuleContext context, QualifiedSelection selection, 
            RemoveParametersRefactoring quickFixRefactoring, RubberduckParserState state)
            : base(context, selection, InspectionsUI.RemoveUnusedParameterQuickFix)
        {
            _quickFixRefactoring = quickFixRefactoring;
            _state = state;
        }

        public override void Fix()
        {
            _quickFixRefactoring.QuickFix(_state, Selection);
        }
    }
}
