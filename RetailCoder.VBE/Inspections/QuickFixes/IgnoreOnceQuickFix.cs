using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    public class IgnoreOnceQuickFix : QuickFixBase
    {
        private readonly string _annotationText;

        public IgnoreOnceQuickFix(ParserRuleContext context, QualifiedSelection selection, string inspectionName) 
            : base(context, selection, InspectionsUI.IgnoreOnce)
        {
            _annotationText = "'" + Annotations.AnnotationMarker +
                              Annotations.IgnoreInspection + ' ' + inspectionName;
        }

        public override bool CanFixInModule { get { return false; } } // not quite "once" if applied to entire module
        public override bool CanFixInProject { get { return false; } } // use "disable this inspection" instead of ignoring across the project

        public override void Fix()
        {
            var module = Selection.QualifiedName.Component.CodeModule;
            {
                var insertLine = Selection.Selection.StartLine;
                while (insertLine != 1 && module.GetLines(insertLine - 1, 1).EndsWith(" _"))
                {
                    insertLine--;
                }
                var codeLine = insertLine == 1 ? string.Empty : module.GetLines(insertLine - 1, 1);
                var annotationText = _annotationText;
                var ignoreAnnotation = "'" + Annotations.AnnotationMarker + Annotations.IgnoreInspection;

                int commentStart;
                if (codeLine.HasComment(out commentStart) && codeLine.Substring(commentStart).StartsWith(ignoreAnnotation))
                {
                    var indentation = codeLine.Length - codeLine.TrimStart().Length;
                    annotationText = string.Format("{0}{1},{2}", 
                                                   new string(' ', indentation), 
                                                   _annotationText,
                                                   codeLine.Substring(indentation + ignoreAnnotation.Length));
                    module.ReplaceLine(insertLine - 1, annotationText);
                }
                else
                {
                    module.InsertLines(insertLine, annotationText);
                }
            }
        }
    }
}
