using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class ParameterNotUsedInspectionResult : CodeInspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public ParameterNotUsedInspectionResult(IInspection inspection, string result,
            ParserRuleContext context, QualifiedMemberName qualifiedName, bool isInterfaceImplementation, 
            RemoveParametersRefactoring refactoring, RubberduckParserState parseResult)
            : base(inspection, result, qualifiedName.QualifiedModuleName, context)
        {
            _quickFixes = isInterfaceImplementation ? new CodeInspectionQuickFix[] {} : new[]
            {
                new RemoveUnusedParameterQuickFix(Context, QualifiedSelection, refactoring, parseResult),
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }
    }

    public class RemoveUnusedParameterQuickFix : CodeInspectionQuickFix
    {
        private readonly RemoveParametersRefactoring _quickFixRefactoring;
        private readonly RubberduckParserState _parseResult;

        public RemoveUnusedParameterQuickFix(ParserRuleContext context, QualifiedSelection selection, 
            RemoveParametersRefactoring quickFixRefactoring, RubberduckParserState parseResult)
            : base(context, selection, RubberduckUI.Inspections_RemoveUnusedParameter)
        {
            _quickFixRefactoring = quickFixRefactoring;
            _parseResult = parseResult;
        }

        public override void Fix()
        {
            _quickFixRefactoring.QuickFix(_parseResult, Selection);
        }
    }
}