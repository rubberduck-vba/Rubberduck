using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class ParameterCanBeByValInspectionResult : InspectionResultBase
    {
        private readonly Declaration _declaration;
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public ParameterCanBeByValInspectionResult(IInspection inspection, Declaration declaration)
            : base(inspection, declaration)
        {
            _declaration = declaration;
            _quickFixes = new[]
            {
                new PassParameterByValueQuickFix(Context, QualifiedSelection), 
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }

        protected override Declaration Target
        {
            get { return _declaration; }
        }

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
            var parameter = Context.GetText();
            var newContent = string.Concat(Tokens.ByVal, " ", parameter.Replace(Tokens.ByRef, string.Empty).Trim());
            var selection = Selection.Selection;

            var module = Selection.QualifiedName.Component.CodeModule;
            var lines = module.get_Lines(selection.StartLine, selection.LineCount);

            var result = lines.Replace(parameter, newContent);
            module.ReplaceLine(selection.StartLine, result);
        }
    }
}