using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class ParameterNotUsedInspectionResult : CodeInspectionResultBase
    {
        private readonly bool _isInterfaceImplementation;
        private readonly RemoveParametersRefactoring _quickFixRefactoring;
        private readonly VBProjectParseResult _parseResult;

        public ParameterNotUsedInspectionResult(string inspection, CodeInspectionSeverity type,
            ParserRuleContext context, QualifiedMemberName qualifiedName, bool isInterfaceImplementation, RemoveParametersRefactoring quickFixRefactoring, VBProjectParseResult parseResult)
            : base(inspection, type, qualifiedName.QualifiedModuleName, context)
        {
            _isInterfaceImplementation = isInterfaceImplementation;
            _quickFixRefactoring = quickFixRefactoring;
            _parseResult = parseResult;
        }

        public override IDictionary<string, Action> GetQuickFixes()
        {
            var result = new Dictionary<string, Action>();
            if (!_isInterfaceImplementation)
            {
                result.Add(RubberduckUI.Inspections_RemoveUnusedParameter, RemoveUnusedParameter);
            }

            return result;
        }

        private void RemoveUnusedParameter()
        {
            _quickFixRefactoring.QuickFix(_parseResult, this.QualifiedSelection);
        }
    }
}