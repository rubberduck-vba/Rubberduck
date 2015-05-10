using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class ParameterNotUsedInspectionResult : CodeInspectionResultBase
    {
        private readonly bool _isInterfaceImplementation;

        public ParameterNotUsedInspectionResult(string inspection, CodeInspectionSeverity type,
            ParserRuleContext context, QualifiedMemberName qualifiedName, bool isInterfaceImplementation)
            : base(inspection, type, qualifiedName.QualifiedModuleName, context)
        {
            _isInterfaceImplementation = isInterfaceImplementation;
        }

        public override IDictionary<string, Action<VBE>> GetQuickFixes()
        {
            var result = new Dictionary<string, Action<VBE>>();
            if (!_isInterfaceImplementation)
            {
                // don't bother implementing this without implementing a ChangeSignatureRefactoring
                //{"Remove unused parameter", RemoveUnusedParameter}
            };

            return result;
        }
    }
}