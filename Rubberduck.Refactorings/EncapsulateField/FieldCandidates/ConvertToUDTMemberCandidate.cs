using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.EncapsulateField
{

    public interface IConvertToUDTMember : IEncapsulatableField
    {
        string UDTMemberDeclaration { get; }
        IObjectStateUDT ObjectStateUDT { set; get; }
        IEncapsulatableField WrappedCandidate { get; }
    }

    public class ConvertToUDTMember : IConvertToUDTMember
    {
        private readonly IEncapsulatableField _wrapped;
        public ConvertToUDTMember(IEncapsulatableField candidate, IObjectStateUDT objStateUDT)
        {
            _wrapped = candidate;
            BackingIdentifier = _wrapped.PropertyIdentifier;
            ObjectStateUDT = objStateUDT;
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

        public IEncapsulatableField WrappedCandidate => _wrapped;

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

        public string BackingIdentifier { get; set; }

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
                    return $"{ObjectStateUDT.FieldIdentifier}.{udtm.Parent.PropertyIdentifier}.{BackingIdentifier}";
                }
                return $"{ObjectStateUDT.FieldIdentifier}.{BackingIdentifier}";
            }
        }

        public string ReferenceAccessor(IdentifierReference idRef)
        {
            if (idRef.QualifiedModuleName != QualifiedModuleName)
            {
                return $"{QualifiedModuleName.ComponentName}.{PropertyIdentifier}";
            }
            return  $"{BackingIdentifier}";
        }

        public string IdentifierName => _wrapped.IdentifierName;

        public QualifiedModuleName QualifiedModuleName => _wrapped.QualifiedModuleName;

        public string AsTypeName => _wrapped.AsTypeName;

        public bool TryValidateEncapsulationAttributes(out string errorMessage)
        {
            return ConflictFinder.TryValidateEncapsulationAttributes(this, out errorMessage);
        }

        public IEnumerable<PropertyAttributeSet> PropertyAttributeSets
        {
            get
            {
                if (_wrapped is IUserDefinedTypeCandidate udt && udt.TypeDeclarationIsPrivate)
                {
                    var sets = new List<PropertyAttributeSet>();
                    foreach (var member in udt.Members)
                    {
                        sets.Add(CreateMemberPropertyAttributeSet(member));
                    }
                    return sets;
                }
                return new List<PropertyAttributeSet>() { AsPropertyAttributeSet };
            }
        }

        private PropertyAttributeSet CreateMemberPropertyAttributeSet (IUserDefinedTypeMemberCandidate udtMember)
        {
            return new PropertyAttributeSet()
            {
                PropertyName = udtMember.PropertyIdentifier,
                BackingField = $"{ObjectStateUDT.FieldIdentifier}.{udtMember.Parent.PropertyIdentifier}.{udtMember.BackingIdentifier}",
                AsTypeName = udtMember.PropertyAsTypeName,
                ParameterName = udtMember.ParameterName,
                GenerateLetter = udtMember.ImplementLet,
                GenerateSetter = udtMember.ImplementSet,
                UsesSetAssignment = udtMember.Declaration.IsObject,
                IsUDTProperty = false //TODO: If udtMember is a UDT, this needs to be true
            };
        }

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
