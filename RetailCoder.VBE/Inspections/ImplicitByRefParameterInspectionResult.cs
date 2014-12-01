using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.VBA.Parser;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public class ImplicitByRefParameterInspectionResult : CodeInspectionResultBase
    {
        public ImplicitByRefParameterInspectionResult(string inspection, Instruction instruction, CodeInspectionSeverity type)
            : base(inspection, instruction, type)
        {
        }

        public override IDictionary<string, Action<VBE>> GetQuickFixes()
        {
            return !Handled
                ? new Dictionary<string, Action<VBE>>
                {
                    {"Pass parameter by reference explicitly.", PassParameterByRef},
                    {"Pass parameter by value.", PassParameterByVal}
                }
                : new Dictionary<string, Action<VBE>>();
        }

        private void PassParameterByRef(VBE vbe)
        {
            throw new NotImplementedException();
        }

        private void PassParameterByVal(VBE vbe)
        {
            throw new NotImplementedException();
        }
    }
}