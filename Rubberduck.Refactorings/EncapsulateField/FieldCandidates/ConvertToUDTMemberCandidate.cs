using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.EncapsulateField
{

    public class ConvertToUDTMember : IConvertToUDTMember
    {
        private readonly IEncapsulatableField _wrapped;
        public ConvertToUDTMember(IEncapsulatableField candidate, IObjectStateUDT objStateUDT)
        {
            _wrapped = candidate;
            ObjectStateUDT = objStateUDT;
        }

        public string UDTMemberIdentifier
        {
            set => _wrapped.PropertyIdentifier = value;
            get => _wrapped.PropertyIdentifier;
        }

        public virtual string UDTMemberDeclaration
        {
            get
            {
                if (_wrapped is IArrayCandidate array)
                {
                   return array.UDTMemberDeclaration;
                }
                return $"{_wrapped.PropertyIdentifier} As {_wrapped.AsTypeName}";
            }
        }

        public IObjectStateUDT ObjectStateUDT { set; get; }

        public string TargetID => _wrapped.TargetID;

        public Declaration Declaration => _wrapped.Declaration;

        public bool EncapsulateFlag
        {
            set => _wrapped.EncapsulateFlag = value;
            get => _wrapped.EncapsulateFlag;
        }

        public string PropertyIdentifier
        {
            set => _wrapped.PropertyIdentifier = value;
            get => _wrapped.PropertyIdentifier;
        }

        public string PropertyAsTypeName
        {
            set => _wrapped.PropertyAsTypeName = value;
            get => _wrapped.PropertyAsTypeName;
        }

        public bool CanBeReadWrite
        {
            set => _wrapped.CanBeReadWrite = value;
            get => _wrapped.CanBeReadWrite;
        }

        public bool ImplementLet => _wrapped.ImplementLet;

        public bool ImplementSet => _wrapped.ImplementSet;

        public bool IsReadOnly
        {
            set => _wrapped.IsReadOnly = value;
            get => _wrapped.IsReadOnly;
        }

        public string ParameterName => _wrapped.ParameterName;

        public IValidateVBAIdentifiers NameValidator
        {
            set => _wrapped.NameValidator = value;
            get => _wrapped.NameValidator;
        }

        public IEncapsulateFieldConflictFinder ConflictFinder
        {
            set => _wrapped.ConflictFinder = value;
            get => _wrapped.ConflictFinder;
        }

        public string AccessorInProperty
        {
            get
            {
                if (_wrapped is IUserDefinedTypeCandidate udt)
                {
                    return $"{ObjectStateUDT.FieldIdentifier}.{udt.PropertyIdentifier}.{UDTMemberIdentifier}";
                }
                return $"{ObjectStateUDT.FieldIdentifier}.{UDTMemberIdentifier}";
            }
        }

        public string AccessorLocalReference
            => $"{ObjectStateUDT.FieldIdentifier}.{UDTMemberIdentifier}";

        public string AccessorExternalReference
            => $"{QualifiedModuleName.ComponentName}.{PropertyIdentifier}";

        public string IdentifierName => _wrapped.IdentifierName;

        public QualifiedModuleName QualifiedModuleName => _wrapped.QualifiedModuleName;

        public string AsTypeName => _wrapped.AsTypeName;

        public bool TryValidateEncapsulationAttributes(out string errorMessage)
        {
            return ConflictFinder.TryValidateEncapsulationAttributes(this, out errorMessage);
        }

        public IEnumerable<IPropertyGeneratorAttributes> PropertyAttributeSets
        {
            get
            {
                //if (TypeDeclarationIsPrivate)
                //{
                    if (_wrapped is IUserDefinedTypeCandidate udt)
                {
                    return _wrapped.PropertyAttributeSets;
                }
                    //var specs = new List<IPropertyGeneratorAttributes>();
                    //foreach (var member in Members)
                    //{
                    //    specs.Add(member.AsPropertyGeneratorSpec);
                    //}
                    //return specs;
                //}
                return new List<IPropertyGeneratorAttributes>() { AsPropertyAttributeSet };
            }
        }


        private IPropertyGeneratorAttributes AsPropertyAttributeSet
        {
            get
            {
                return new PropertyAttributeSet()
                {
                    PropertyName = PropertyIdentifier,
                    BackingField = AccessorInProperty, // ReferenceWithinNewProperty,
                    AsTypeName = PropertyAsTypeName,
                    ParameterName = ParameterName,
                    GenerateLetter = ImplementLet,
                    GenerateSetter = ImplementSet,
                    UsesSetAssignment = Declaration.IsObject,
                    IsUDTProperty = true
                };
            }
        }
    }
}
