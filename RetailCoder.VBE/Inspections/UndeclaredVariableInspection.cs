using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class UndeclaredVariableInspection : InspectionBase
    {
        public UndeclaredVariableInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Error)
        {
        }

        public override string Meta { get { return InspectionsUI.UndeclaredVariableInspectionMeta; } }
        public override string Description { get { return InspectionsUI.UndeclaredVariableInspectionResultFormat; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            return UserDeclarations.Where(item => item.IsUndeclared)
                .Select(item => new UndeclaredVariableInspectionResult(this, item));
        }
    }

    public class UndeclaredVariableInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public UndeclaredVariableInspectionResult(IInspection inspection, Declaration target)
            : base(inspection, target)
        {
            _quickFixes = new CodeInspectionQuickFix[]
            {
                new IntroduceLocalVariableQuickFix(target), 
                new IgnoreOnceQuickFix(target.Context, target.QualifiedSelection, inspection.Name), 
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes
        {
            get { return _quickFixes; }
        }

        public override string Description { get { return string.Format(InspectionsUI.UndeclaredVariableInspectionResultFormat, Target.IdentifierName).Captialize(); } }
    }

    public class IntroduceLocalVariableQuickFix : CodeInspectionQuickFix
    {
        private readonly Declaration _undeclared;

        public IntroduceLocalVariableQuickFix(Declaration undeclared) 
            : base(undeclared.Context, undeclared.QualifiedSelection, InspectionsUI.IntroduceLocalVariableQuickFix)
        {
            _undeclared = undeclared;
        }

        public override bool CanFixInModule { get { return true; } }
        public override bool CanFixInProject { get { return true; } }

        public override void Fix()
        {
            var instruction = Tokens.Dim + ' ' + _undeclared.IdentifierName + ' ' + Tokens.As + ' ' + Tokens.Variant;
            Selection.QualifiedName.Component.CodeModule.InsertLines(Selection.Selection.StartLine, instruction);
        }
    }
}