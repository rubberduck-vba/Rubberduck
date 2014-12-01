using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.VBA.Parser;
using Rubberduck.VBA.Parser.Grammar;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public class OptionExplicitInspectionResult : CodeInspectionResultBase
    {
        public OptionExplicitInspectionResult(string inspection, Instruction instruction, CodeInspectionSeverity type) 
            : base(inspection, instruction, type)
        {
        }

        public override IDictionary<string, Action<VBE>> GetQuickFixes()
        {
            return !Handled
                ? new Dictionary<string, Action<VBE>>
                    {
                        {"Specify Option Explicit.", SpecifyOptionExplicit}
                    }
                : new Dictionary<string, Action<VBE>>();
        }

        private void SpecifyOptionExplicit(VBE vbe)
        {
            var modules = vbe.FindCodeModules(Instruction.Line.ProjectName, Instruction.Line.ComponentName);
            foreach (var codeModule in modules)
            {
                codeModule.InsertLines(1, string.Concat(ReservedKeywords.Option, " ", ReservedKeywords.Explicit));
            }

            Handled = true;
        }
    }
}