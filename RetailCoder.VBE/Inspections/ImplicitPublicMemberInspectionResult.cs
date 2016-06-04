using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class ImplicitPublicMemberInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public ImplicitPublicMemberInspectionResult(IInspection inspection, QualifiedContext<ParserRuleContext> qualifiedContext, Declaration item)
            : base(inspection, item)
        {
            _quickFixes = new CodeInspectionQuickFix[]
            {
                new SpecifyExplicitPublicModifierQuickFix(qualifiedContext.Context, QualifiedSelection), 
                new IgnoreOnceQuickFix(qualifiedContext.Context, QualifiedSelection, Inspection.AnnotationName), 
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get
            {
                return string.Format(InspectionsUI.ImplicitPublicMemberInspectionResultFormat, Target.IdentifierName);
            }
        }

        public override NavigateCodeEventArgs GetNavigationArgs()
        {
            return new NavigateCodeEventArgs(Target);
        }
    }

    public class SpecifyExplicitPublicModifierQuickFix : CodeInspectionQuickFix
    {
        public SpecifyExplicitPublicModifierQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, InspectionsUI.SpecifyExplicitPublicModifierQuickFix)
        {
        }

        public override void Fix()
        {
            var selection = Context.GetSelection();
            var module = Selection.QualifiedName.Component.CodeModule;

            var signatureLine = selection.StartLine;

            var oldContent = module.get_Lines(signatureLine, 1);
            var newContent = Tokens.Public + ' ' + oldContent;

            module.DeleteLines(signatureLine);
            module.InsertLines(signatureLine, newContent);
        }
    }
}
