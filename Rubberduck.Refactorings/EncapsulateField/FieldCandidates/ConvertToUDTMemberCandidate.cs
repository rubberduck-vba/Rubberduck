using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.EncapsulateField
{

    public interface IConvertToUDTMember : IEncapsulateFieldCandidate
    {
        string UDTMemberDeclaration { get; }
        IEncapsulateFieldCandidate WrappedCandidate { get; }
        IObjectStateUDT ObjectStateUDT { set; get; }
    }

    public class ConvertToUDTMember : IConvertToUDTMember
    {
        private int _hashCode;
        private readonly string _uniqueID;
        private readonly IEncapsulateFieldCandidate _wrapped;
        public ConvertToUDTMember(IEncapsulateFieldCandidate candidate, IObjectStateUDT objStateUDT)
        {
            _wrapped = candidate;
            PropertyIdentifier = _wrapped.PropertyIdentifier;
            ObjectStateUDT = objStateUDT;
            _uniqueID = BuildUniqueID(candidate, objStateUDT);
            _hashCode = _uniqueID.GetHashCode();
        }

        public virtual string UDTMemberDeclaration
        {
            get
            {
                if (_wrapped is IArrayCandidate array)
                {
                   return array.UDTMemberDeclaration;
                }
                return $"{BackingIdentifier} As {_wrapped.AsTypeName}";
            }
        }

        public IEncapsulateFieldCandidate WrappedCandidate => _wrapped;

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

        public string PropertyAsTypeName => _wrapped.PropertyAsTypeName;

        public string BackingIdentifier
        {
            set { }
            get => PropertyIdentifier;
        }
        public string BackingAsTypeName => Declaration.AsTypeName;

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

        private string AccessorInProperty
        {
            get
            {
                if (_wrapped is IUserDefinedTypeMemberCandidate udtm)
                {
                    return $"{ObjectStateUDT.FieldIdentifier}.{udtm.UDTField.PropertyIdentifier}.{BackingIdentifier}";
                }
                return $"{ObjectStateUDT.FieldIdentifier}.{BackingIdentifier}";
            }
        }

        public string IdentifierForReference(IdentifierReference idRef)
        {
            if (idRef.QualifiedModuleName != QualifiedModuleName)
            {
                return PropertyIdentifier;
            }
            return  BackingIdentifier;
        }

        public string IdentifierName => _wrapped.IdentifierName;

        public QualifiedModuleName QualifiedModuleName => _wrapped.QualifiedModuleName;

        public string AsTypeName => _wrapped.AsTypeName;

        public bool TryValidateEncapsulationAttributes(out string errorMessage)
        {
            errorMessage = string.Empty;
            if (!_wrapped.EncapsulateFlag) { return true; }

            if (_wrapped is IArrayCandidate ac)
            {
                if (ac.HasExternalRedimOperation(out errorMessage))
                {
                    return false;
                }
            }
            return ConflictFinder.TryValidateEncapsulationAttributes(this, out errorMessage);
        }

        public IEnumerable<PropertyAttributeSet> PropertyAttributeSets
        {
            get
            {
                var modifiedSets = new List<PropertyAttributeSet>();
                var sets = _wrapped.PropertyAttributeSets;
                for (var idx = 0; idx < sets.Count(); idx++)
                {
                    var attributeSet = sets.ElementAt(idx);
                    var fields = attributeSet.BackingField.Split(new char[] { '.' });

                    attributeSet.BackingField = fields.Count() > 1
                        ? $"{ObjectStateUDT.FieldIdentifier}.{attributeSet.BackingField.CapitalizeFirstLetter()}"
                        : $"{ObjectStateUDT.FieldIdentifier}.{attributeSet.PropertyName.CapitalizeFirstLetter()}";

                    modifiedSets.Add(attributeSet);
                }
                return modifiedSets;
            }
        }

        public override bool Equals(object obj)
        {
            return obj != null
                && obj is ConvertToUDTMember convertWrapper
                && BuildUniqueID(convertWrapper, convertWrapper.ObjectStateUDT) == _uniqueID;
        }

        public override int GetHashCode() => _hashCode;

        private static string BuildUniqueID(IEncapsulateFieldCandidate candidate, IObjectStateUDT field) 
            => $"{candidate.QualifiedModuleName.Name}.{field.IdentifierName}.{candidate.IdentifierName}";

        private PropertyAttributeSet AsPropertyAttributeSet
        {
            get
            {
                return new PropertyAttributeSet()
                {
                    PropertyName = PropertyIdentifier,
                    BackingField = AccessorInProperty,
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
