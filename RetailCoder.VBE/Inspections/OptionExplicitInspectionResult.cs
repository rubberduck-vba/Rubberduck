using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Nodes;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class OptionExplicitInspectionResult : CodeInspectionResultBase
    {
        public OptionExplicitInspectionResult(string inspection, CodeInspectionSeverity type, QualifiedModuleName qualifiedName) 
            : base(inspection, type, new CommentNode("", new QualifiedSelection(qualifiedName, Selection.Home)))
        {
        }

        public override IDictionary<string, Action<VBE>> GetQuickFixes()
        {
            return
                new Dictionary<string, Action<VBE>>
                {
                    {"Specify Option Explicit", SpecifyOptionExplicit}
                };
        }

        private void SpecifyOptionExplicit(VBE vbe)
        {
            var module = QualifiedName.Component.CodeModule;
            if (module == null)
            {
                return;
            }

            module.InsertLines(1, Tokens.Option + ' ' + Tokens.Explicit + "\n");
        }
    }
}