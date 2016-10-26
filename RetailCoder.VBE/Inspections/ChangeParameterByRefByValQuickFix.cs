using Antlr4.Runtime;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class ChangeParameterByRefByValQuickFix : CodeInspectionQuickFix
    {
        private readonly string _newToken;

        public ChangeParameterByRefByValQuickFix(ParserRuleContext context, QualifiedSelection selection, string description, string newToken) 
            : base(context, selection, description)
        {
            _newToken = newToken;
        }

        public override void Fix()
        {
            var parameter = Context.GetText();
            var newContent = string.Concat(_newToken, " ", parameter);
            var selection = Selection.Selection;

            var module = Selection.QualifiedName.Component.CodeModule;
            {
                var lines = module.GetLines(selection.StartLine, selection.LineCount);
                var result = lines.Replace(parameter, newContent);
                module.ReplaceLine(selection.StartLine, result);
            }
        }
    }
}