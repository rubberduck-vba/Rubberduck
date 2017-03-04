using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.Results
{
    public class MemberNotOnInterfaceInspectionResult : InspectionResultBase
    {
        private readonly Declaration _asTypeDeclaration;
        private IEnumerable<QuickFixBase> _quickFixes;

        public MemberNotOnInterfaceInspectionResult(IInspection inspection, Declaration target, Declaration asTypeDeclaration)
            : base(inspection, target)
        {
            _asTypeDeclaration = asTypeDeclaration;
        }

        public override IEnumerable<QuickFixBase> QuickFixes
        {
            get
            {
                return _quickFixes ?? (_quickFixes = new QuickFixBase[]
                {
                    new IgnoreOnceQuickFix(Target.Context, QualifiedSelection, Inspection.AnnotationName)
                });
            }
        }

        public override string Description
        {
            get { return string.Format(InspectionsUI.MemberNotOnInterfaceInspectionResultFormat, Target.IdentifierName, _asTypeDeclaration.IdentifierName); }
        }
    }
}
