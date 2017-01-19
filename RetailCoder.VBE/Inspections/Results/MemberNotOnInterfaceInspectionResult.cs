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
        private readonly ParserRuleContext _member;
        private readonly Declaration _asTypeDeclaration;
        private IEnumerable<QuickFixBase> _quickFixes;

        public MemberNotOnInterfaceInspectionResult(Parsing.Symbols.IInspection inspection, Declaration target, ParserRuleContext member, Declaration asTypeDeclaration)
            : base(inspection, target)
        {
            _member = member;
            _asTypeDeclaration = asTypeDeclaration;
        }

        public override IEnumerable<QuickFixBase> QuickFixes
        {
            get
            {
                return _quickFixes ?? (_quickFixes = new QuickFixBase[]
                {
                    new IgnoreOnceQuickFix(_member, QualifiedSelection, Inspection.AnnotationName)
                });
            }
        }

        public override string Description
        {
            get { return string.Format(InspectionsUI.MemberNotOnInterfaceInspectionResultFormat, _member.GetText(), _asTypeDeclaration.IdentifierName); }
        }
    }
}
