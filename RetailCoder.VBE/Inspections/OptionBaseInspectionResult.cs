using System;
using System.Collections.Generic;
using Rubberduck.Parsing.Nodes;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class OptionBaseInspectionResult : CodeInspectionResultBase
    {
        public OptionBaseInspectionResult(string inspection, CodeInspectionSeverity type, QualifiedModuleName qualifiedName)
            : base(inspection, type, new CommentNode("", new QualifiedSelection(qualifiedName, Selection.Home)))
        {
        }

        public override IDictionary<string, Action> GetQuickFixes()
        {
            // removing or changing option statement could break the code.
            return new Dictionary<string, Action>(); 
        }
    }
}