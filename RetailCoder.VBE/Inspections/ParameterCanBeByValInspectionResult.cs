using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class ParameterCanBeByValInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public ParameterCanBeByValInspectionResult(IInspection inspection, Declaration target, ParserRuleContext context, QualifiedMemberName qualifiedName)
            : base(inspection, qualifiedName.QualifiedModuleName, context, target)
        {
            _quickFixes = new[]
            {
                new PassParameterByValueQuickFix(Context, QualifiedSelection), 
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get { return string.Format(InspectionsUI.ParameterCanBeByValInspectionResultFormat, Target.IdentifierName); }
        }
    }

    public class PassParameterByValueQuickFix : CodeInspectionQuickFix
    {
        public PassParameterByValueQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, InspectionsUI.PassParameterByValueQuickFix)
        {
        }

        public override void Fix()
        {
            var selection = Selection.Selection;
            var selectionLength = ((VBAParser.ArgContext) Context).BYREF() == null ? 0 : 6;

            var module = Selection.QualifiedName.Component.CodeModule;
            var lines = module.Lines[selection.StartLine, 1];

            var result = lines.Remove(selection.StartColumn - 1, selectionLength).Insert(selection.StartColumn - 1, Tokens.ByVal + ' ');
            module.ReplaceLine(selection.StartLine, result);
        }
    }
}
