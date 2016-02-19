using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class IdentifierNotUsedInspectionResult : InspectionResultBase
    {
        private readonly Declaration _target;
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public IdentifierNotUsedInspectionResult(IInspection inspection, Declaration target,
            ParserRuleContext context, QualifiedModuleName qualifiedName)
            : base(inspection, qualifiedName, context)
        {
            _target = target;
            _quickFixes = new CodeInspectionQuickFix[]
            {
                new RemoveUnusedDeclarationQuickFix(context, QualifiedSelection), 
                new IgnoreOnceQuickFix(context, QualifiedSelection, Inspection.AnnotationName), 
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get
            {
                return string.Format(InspectionsUI.IdentifierNotUsedInspectionResultFormat, _target.DeclarationType.ToLocalizedString(), _target.IdentifierName);
            }
        }
    }

    /// <summary>
    /// A code inspection quickfix that removes an unused identifier declaration.
    /// </summary>
    public class RemoveUnusedDeclarationQuickFix : CodeInspectionQuickFix
    {
        public RemoveUnusedDeclarationQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, InspectionsUI.RemoveUnusedDeclarationQuickFix)
        {
        }

        public override void Fix()
        {
            var module = Selection.QualifiedName.Component.CodeModule;
            var selection = Selection.Selection;

            var originalCodeLines = module.get_Lines(selection.StartLine, selection.LineCount)
                .Replace("\r\n", " ")
                .Replace("_", string.Empty);

            var originalInstruction = Context.GetText();
            module.DeleteLines(selection.StartLine, selection.LineCount);

            var newInstruction = string.Empty;
            var newCodeLines = string.IsNullOrEmpty(newInstruction)
                ? string.Empty
                : originalCodeLines.Replace(originalInstruction, newInstruction);

            if (!string.IsNullOrEmpty(newCodeLines))
            {
                module.InsertLines(selection.StartLine, newCodeLines);
            }
        }
    }
}