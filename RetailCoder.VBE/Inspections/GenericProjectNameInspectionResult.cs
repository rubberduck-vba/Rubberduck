using System;
using System.Collections.Generic;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Nodes;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class GenericProjectNameInspectionResult : CodeInspectionResultBase
    {
        public GenericProjectNameInspectionResult(string inspection, CodeInspectionSeverity type, QualifiedModuleName qualifiedName) 
            : base(inspection, type, new CommentNode("", new QualifiedSelection(qualifiedName, Selection.Home)))
        {
        }

        public override IDictionary<string, Action> GetQuickFixes()
        {
            return new Dictionary<string, Action>();
        }
    }
}
