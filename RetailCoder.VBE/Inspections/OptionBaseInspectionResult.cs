using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class OptionBaseInspectionResult : InspectionResultBase
    {
        public OptionBaseInspectionResult(IInspection inspection, QualifiedModuleName qualifiedName)
            : base(inspection, new CommentNode(string.Empty, Tokens.CommentMarker, new QualifiedSelection(qualifiedName, Selection.Home)))
        {
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes
        {
            get
            {
                return new CodeInspectionQuickFix[]
                {
                    new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName)
                };
            }
        }

        public override string Description
        {
            get { return string.Format(Inspection.Description, QualifiedName.ComponentName); }
        }
    }
}
