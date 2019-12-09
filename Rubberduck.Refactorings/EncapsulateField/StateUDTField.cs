using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class StateUDTField : EncapsulateFieldCandidate, IEncapsulatedUserDefinedTypeField
    {
        public StateUDTField(string identifier, string asTypeName, QualifiedModuleName qmn, IEncapsulateFieldNamesValidator validator)
            : base(identifier, asTypeName, qmn, validator)
        {
            PropertyName = identifier;
            NewFieldName = identifier;
            AsTypeName = asTypeName;
        }

        public override string PropertyName { set; get; }

        public override string NewFieldName { set; get; }

        public override bool IsSelfConsistent => _validator.IsValidVBAIdentifier(PropertyName, DeclarationType.Property)
                            && !(PropertyName.EqualsVBAIdentifier(ParameterName)
                                    || PropertyName.EqualsVBAIdentifier(ParameterName));

        public IEnumerable<IEncapsulatedUserDefinedTypeMember> Members { get; }
        public void AddMember(IEncapsulatedUserDefinedTypeMember member) { throw new NotImplementedException(); }
        public bool FieldQualifyMemberPropertyNames { set; get; }
        public bool TypeDeclarationIsPrivate { set; get; } = true;
    }
}
