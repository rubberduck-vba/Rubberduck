using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    public class ChangeDimToPrivateQuickFix : QuickFixBase
    {
        public ChangeDimToPrivateQuickFix(ParserRuleContext context, QualifiedSelection selection) 
            : base(context, selection, InspectionsUI.ChangeDimToPrivateQuickFix)
        {
        }

        public override bool CanFixInModule { get { return true; } }
        public override bool CanFixInProject { get { return true; } }

        public override void Fix()
        {
            var module = Selection.QualifiedName.Component.CodeModule;
            if (module == null)
            {
                return;
            }

            var context = (VBAParser.VariableStmtContext)Context.Parent.Parent;
            var newInstruction = Tokens.Private + " ";
            for (var i = 1; i < context.ChildCount; i++)
            {
                // index 0 would be the 'Dim' keyword
                newInstruction += context.GetChild(i).GetText();
            }

            var selection = context.GetSelection();
            var oldContent = module.GetLines(selection);
            var newContent = oldContent.Replace(context.GetText(), newInstruction);

            module.DeleteLines(selection);
            module.InsertLines(selection.StartLine, newContent);
        }
    }
}