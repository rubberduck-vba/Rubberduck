using Antlr4.Runtime;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class IgnoreOnceQuickFix : CodeInspectionQuickFix
    {
        private readonly string _annotationText;
        private readonly string _inspectionName;

        public IgnoreOnceQuickFix(ParserRuleContext context, QualifiedSelection selection, string inspectionName) 
            : base(context, selection, InspectionsUI.IgnoreOnce)
        {
            _inspectionName = inspectionName;
            _annotationText = "'" + Parsing.Grammar.Annotations.AnnotationMarker +
                              Parsing.Grammar.Annotations.IgnoreInspection + ' ' + inspectionName;
        }

        public override bool CanFixInModule { get { return false; } } // not quite "once" if applied to entire module
        public override bool CanFixInProject { get { return false; } } // use "disable this inspection" instead of ignoring across the project

        public override void Fix()
        {
            var codeModule = Selection.QualifiedName.Component.CodeModule;
            var insertLine = Selection.Selection.StartLine;

            var codeLine = insertLine == 1 ? string.Empty : codeModule.get_Lines(insertLine - 1, 1);
            var annotationText = _annotationText;
            var ignoreAnnotation = "'" + Parsing.Grammar.Annotations.AnnotationMarker + Parsing.Grammar.Annotations.IgnoreInspection;

            int commentStart;
            if (codeLine.HasComment(out commentStart) && codeLine.Substring(commentStart).StartsWith(ignoreAnnotation))
            {
                annotationText = codeLine + ' ' + _inspectionName;
                codeModule.ReplaceLine(insertLine - 1, annotationText);
            }
            else
            {
                codeModule.InsertLines(insertLine, annotationText);
            }
        }
    }
}