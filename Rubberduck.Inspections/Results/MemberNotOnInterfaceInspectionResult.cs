using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.Results
{
    public class MemberNotOnInterfaceInspectionResult : InspectionResultBase
    {
        private readonly Declaration _asTypeDeclaration;

        public MemberNotOnInterfaceInspectionResult(IInspection inspection, Declaration target, Declaration asTypeDeclaration)
            : base(inspection, target)
        {
            _asTypeDeclaration = asTypeDeclaration;
        }

        public override string Description
        {
            get { return string.Format(InspectionsUI.MemberNotOnInterfaceInspectionResultFormat, Target.IdentifierName, _asTypeDeclaration.IdentifierName); }
        }
    }
}
