using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.RemoveParameters;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class ParameterNotUsedInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public ParameterNotUsedInspectionResult(IInspection inspection, string result,
            ParserRuleContext context, QualifiedMemberName qualifiedName, bool isInterfaceImplementation, 
            RemoveParametersRefactoring refactoring, RubberduckParserState parseResult)
            : base(inspection, qualifiedName.QualifiedModuleName, context)
        {
            _quickFixes = isInterfaceImplementation ? new CodeInspectionQuickFix[] {} : new CodeInspectionQuickFix[]
            {
                new RemoveUnusedParameterQuickFix(Context, QualifiedSelection, refactoring, parseResult),
                new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName), 
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get { return string.Format(InspectionsUI.ParameterNotUsedInspectionResultFormat, Target.IdentifierName); }
        }
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