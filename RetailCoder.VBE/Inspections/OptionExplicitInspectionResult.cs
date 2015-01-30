using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Extensions;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public class OptionExplicitInspectionResult : CodeInspectionResultBase
    {
        public OptionExplicitInspectionResult(string inspection, ParserRuleContext context, CodeInspectionSeverity type, string project, string module) 
            : base(inspection, context, type, project, module)
        {
            _project = project;
            _module = module;
        }

        private readonly string _project;
        private readonly string _module;

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
            var modules = vbe.FindCodeModules(_project, _module);
            foreach (var codeModule in modules)
            {
                codeModule.InsertLines(1, ReservedKeywords.Option + " " + ReservedKeywords.Explicit);
            }
        }
    }
}