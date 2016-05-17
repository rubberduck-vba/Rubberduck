using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Nodes;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class OptionBaseInspectionResult : InspectionResultBase
    {
        public OptionBaseInspectionResult(IInspection inspection, QualifiedModuleName qualifiedName)
            : base(inspection, new CommentNode(string.Empty, Tokens.CommentMarker, new QualifiedSelection(qualifiedName, Selection.Home)))
        {
        }

        public override string Description
        {
            get { return string.Format(Inspection.Description, QualifiedName.ComponentName); }
        }
    }
}
