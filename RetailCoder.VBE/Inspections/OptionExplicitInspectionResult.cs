using System;
using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Nodes;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Inspections
{
    public class OptionExplicitInspectionResult : CodeInspectionResultBase
    {
        public OptionExplicitInspectionResult(string inspection, CodeInspectionSeverity type, QualifiedModuleName qualifiedName, IRubberduckFactory<IRubberduckCodePane> factory) 
            : base(inspection, type, new CommentNode("", new QualifiedSelection(qualifiedName, Selection.Home, factory)))
        {
        }

        public override IDictionary<string, Action> GetQuickFixes()
        {
            return
                new Dictionary<string, Action>
                {
                    {"Specify Option Explicit", SpecifyOptionExplicit}
                };
        }

        private void SpecifyOptionExplicit()
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