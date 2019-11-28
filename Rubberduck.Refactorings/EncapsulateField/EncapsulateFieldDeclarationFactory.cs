using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldDeclarationFactory
    {
        public static IEncapsulateFieldCandidate EncapsulateDeclaration(Declaration target, IEncapsulateFieldNamesValidator validator)
        {
            var candidate = new EncapsulateFieldCandidate(target, validator);

            if (target.IsArray)
            {
                candidate.EncapsulationAttributes.ImplementLetSetterType = false;
                candidate.EncapsulationAttributes.ImplementSetSetterType = false;
                candidate.EncapsulationAttributes.AsTypeName = Tokens.Variant;
                candidate.CanBeReadWrite = false;
                candidate.IsReadOnly = true;
                return candidate;
            }
            else if (target.AsTypeName.Equals(Tokens.Variant))
            {
                candidate.EncapsulationAttributes.ImplementLetSetterType = true;
                candidate.EncapsulationAttributes.ImplementSetSetterType = true;
                return candidate;
            }
            else if (target.IsObject)
            {
                candidate.EncapsulationAttributes.ImplementLetSetterType = false;
                candidate.EncapsulationAttributes.ImplementSetSetterType = true;
                return candidate;
            }
            else if (target.AsTypeDeclaration?.DeclarationType.Equals(DeclarationType.UserDefinedType) ?? false)
            {
                candidate.EncapsulationAttributes.ImplementLetSetterType = true;
                candidate.EncapsulationAttributes.ImplementSetSetterType = false;
                return candidate;
            }
            else if (target.DeclarationType.Equals(DeclarationType.UserDefinedTypeMember))
            {
                //TODO: This may need to pass back thru using it's AsTypeName
                candidate.EncapsulationAttributes.ImplementLetSetterType = true;
                candidate.EncapsulationAttributes.ImplementSetSetterType = false;
                return candidate;
            }

            candidate.EncapsulationAttributes.ImplementLetSetterType = true;
            candidate.EncapsulationAttributes.ImplementSetSetterType = false;
            return candidate;
        }

        //public static IEncapsulatedFieldDeclaration EncapsulateDeclaration(Declaration udtMember, IEncapsulatedFieldDeclaration instantiatingField, IEncapsulateFieldNamesValidator validator)
        //{
        //    var targetIDPair = new KeyValuePair<Declaration, string>(udtMember, $"{instantiatingField.Declaration.IdentifierName}.{udtMember.IdentifierName}");
        //    var encapsulatedUdtMember = EncapsulateDeclaration(udtMember, validator);
        //    return EncapsulatedUserDefinedTypeMember.Decorate(encapsulatedUdtMember, instantiatingField, false); // HasMultipleInstantiationsOfSameType(instantiatingField.Declaration, targetIDPair));
        //}

        //    private bool HasMultipleInstantiationsOfSameType(Declaration udtVariable, KeyValuePair<Declaration, string> targetIDPair)
        //    {
        //        var udt = _udtFieldToUdtDeclarationMap[udtVariable].Item1;
        //        var otherVariableOfTheSameType = _udtFieldToUdtDeclarationMap.Keys.Where(k => k != udtVariable && _udtFieldToUdtDeclarationMap[k].Item1 == udt);
        //        return otherVariableOfTheSameType.Any();
        //    }
    }
}
