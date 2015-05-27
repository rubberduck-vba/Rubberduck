using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.VBEditor;
using Rubberduck.UI;

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

        public override IDictionary<string, Action> GetQuickFixes()
        {
            var result = new Dictionary<string, Action>();
            if (!_isInterfaceImplementation)
            {
                // todo: use RemoveParameter refactoring
                //{RubberduckUI.InspectionsRemoveUnusedParameter, RemoveUnusedParameter}
            };

            return result;
        }
    }
}