using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    public class ApplicationWorksheetFunctionQuickFix : QuickFixBase
    {
        private readonly string _memberName;

        public ApplicationWorksheetFunctionQuickFix(QualifiedSelection selection, string memberName)
            : base(null, selection, InspectionsUI.ApplicationWorksheetFunctionQuickFix)
        {
            _memberName = memberName;
        }

        public override bool CanFixInModule { get { return true; } }
        public override bool CanFixInProject { get { return true; } }

        public override void Fix()
        {
            var module = Selection.QualifiedName.Component.CodeModule;
            
            var oldContent = module.GetLines(Selection.Selection);
            var newCall = string.Format("WorksheetFunction.{0}", _memberName);
            var start = Selection.Selection.StartColumn - 1;
            //The member being called will always be a single token, so this will always be safe (it will be a single line).
            var end = Selection.Selection.EndColumn - 1;
            var newContent = oldContent.Substring(0, start) + newCall + 
                (oldContent.Length > end
                ? oldContent.Substring(end, oldContent.Length - end)
                : string.Empty);

            module.DeleteLines(Selection.Selection);
            module.InsertLines(Selection.Selection.StartLine, newContent);
        }
    }
}
