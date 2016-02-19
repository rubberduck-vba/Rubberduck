using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class AssignedByValParameterInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public AssignedByValParameterInspectionResult(IInspection inspection, ParserRuleContext context, QualifiedMemberName qualifiedName)
            : base(inspection, qualifiedName.QualifiedModuleName, context)
        {
            _quickFixes = new[]
            {
                new PassParameterByReferenceQuickFix(context, QualifiedSelection),
            };
        }

        public override string Description
        {
            get
            {
                return string.Format(InspectionsUI.AssignedByValParameterInspectionResultFormat, Target.IdentifierName);
            }
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }
    }

    /// <summary>
    /// Encapsulates a code inspection quickfix that changes a ByVal parameter into an explicit ByRef parameter.
    /// </summary>
    public class PassParameterByReferenceQuickFix : CodeInspectionQuickFix
    {
        public PassParameterByReferenceQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, InspectionsUI.PassParameterByReferenceQuickFix)
        {
        }

        public override void Fix()
        {
            var parameter = Context.GetText();
            var newContent = string.Concat(Tokens.ByRef, " ", parameter.Replace(Tokens.ByVal, string.Empty).Trim());
            var selection = Selection.Selection;

            var module = Selection.QualifiedName.Component.CodeModule;
            var lines = module.get_Lines(selection.StartLine, selection.LineCount);

            var result = lines.Replace(parameter, newContent);
            module.ReplaceLine(selection.StartLine, result);
        }
    }
}