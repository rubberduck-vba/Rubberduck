using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.PostProcessing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Results
{
    public class IdentifierNotUsedInspectionResult : InspectionResultBase
    {
        private IEnumerable<IQuickFix> _quickFixes;
        private readonly ParserRuleContext _context;
        private readonly IModuleRewriter _rewriter;

        public IdentifierNotUsedInspectionResult(IInspection inspection, Declaration target,
            ParserRuleContext context, QualifiedModuleName qualifiedName, IModuleRewriter rewriter)
            : base(inspection, qualifiedName, context, target)
        {
            _context = context;
            _rewriter = rewriter;
        }

        public override IEnumerable<IQuickFix> QuickFixes
        {
            get
            {
                return _quickFixes ?? (_quickFixes = new IQuickFix[]
                {
                    new RemoveUnusedDeclarationQuickFix(_context, QualifiedSelection, Target, _rewriter), 
                    new IgnoreOnceQuickFix(_context, QualifiedSelection, Inspection.AnnotationName)
                });
            }
        }

        public override string Description 
        {
            get
            {
                return string.Format(InspectionsUI.IdentifierNotUsedInspectionResultFormat, Target.DeclarationType.ToLocalizedString(), Target.IdentifierName).Captialize();
            }
        }

        public override NavigateCodeEventArgs GetNavigationArgs()
        {
            return new NavigateCodeEventArgs(Target);
        }
    }
}
