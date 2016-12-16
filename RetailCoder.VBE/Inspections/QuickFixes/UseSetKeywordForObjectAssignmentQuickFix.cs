using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class UseSetKeywordForObjectAssignmentQuickFix : QuickFixBase
    {
        public UseSetKeywordForObjectAssignmentQuickFix(IdentifierReference reference)
            : base(context: reference.Context.Parent.Parent as ParserRuleContext, // ImplicitCallStmt_InStmtContext 
                selection: new QualifiedSelection(reference.QualifiedModuleName, reference.Selection),
                description: InspectionsUI.SetObjectVariableQuickFix)
        {
        }

        public override bool CanFixInModule { get { return true; } }
        public override bool CanFixInProject { get { return true; } }

        public override void Fix()
        {
            var module = Selection.QualifiedName.Component.CodeModule;
            {
                var codeLine = module.GetLines(Selection.Selection.StartLine, 1);

                var letStatementLeftSide = Context.GetText();
                var setStatementLeftSide = Tokens.Set + ' ' + letStatementLeftSide;

                var correctLine = codeLine.Replace(letStatementLeftSide, setStatementLeftSide);
                module.ReplaceLine(Selection.Selection.StartLine, correctLine);
            }
        }
    }
}