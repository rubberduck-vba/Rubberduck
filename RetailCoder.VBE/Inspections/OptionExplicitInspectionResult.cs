using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public class OptionExplicitInspectionResult : CodeInspectionResultBase
    {
        public OptionExplicitInspectionResult(string inspection, SyntaxTreeNode node, CodeInspectionSeverity type) 
            : base(inspection, node, type)
        {
        }

        public override IDictionary<string, Action<VBE>> GetQuickFixes()
        {
            return !Handled
                ? new Dictionary<string, Action<VBE>>
                    {
                        {"Specify Option Explicit", SpecifyOptionExplicit}
                    }
                : new Dictionary<string, Action<VBE>>();
        }

        private void SpecifyOptionExplicit(VBE vbe)
        {
            var instruction = Node.Instruction;
            var modules = vbe.FindCodeModules(instruction.Line.ProjectName, instruction.Line.ComponentName);
            foreach (var codeModule in modules)
            {
                codeModule.InsertLines(1, string.Concat(ReservedKeywords.Option, " ", ReservedKeywords.Explicit));
            }

            Handled = true;
        }
    }
}