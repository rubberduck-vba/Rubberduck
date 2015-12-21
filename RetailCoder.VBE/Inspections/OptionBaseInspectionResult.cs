using System;
using System.Collections.Generic;
using Rubberduck.Parsing.Nodes;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class OptionBaseInspectionResult : CodeInspectionResultBase
    {
        public OptionBaseInspectionResult(IInspection inspection, QualifiedModuleName qualifiedName)
            : base(inspection, inspection.Description, new CommentNode(string.Empty, new QualifiedSelection(qualifiedName, Selection.Home)))
        {
        }
    }
}