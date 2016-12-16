using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.QuickFixes
{
    public class IntroduceLocalVariableQuickFix : QuickFixBase
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