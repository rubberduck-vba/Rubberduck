using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.EncapsulateField.Extensions;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldRefactoringElement
    {
        string IdentifierName { get; }
        QualifiedModuleName QualifiedModuleName { get; }
        string AsTypeName { get; }
    }

    public interface IEncapsulateFieldCandidate : IUsingBackingField
    {
        //IEnumerable<IPropertyGeneratorAttributes> PropertyAttributeSets { get; }
        IEnumerable<KeyValuePair<IdentifierReference, (ParserRuleContext, string)>> ReferenceReplacements { get; }
        string ReferenceQualifier { set; get; }
        void LoadFieldReferenceContextReplacements(string referenceQualifier = null);
        bool ConvertFieldToUDTMember { set; get; }
    }

    public enum AccessorMember { Field, Property }

    public interface IEncapsulateFieldCandidateValidations
    {
        bool HasConflictingPropertyIdentifier { get; }
        bool HasConflictingFieldIdentifier { get; }
    }

    //public class ConvertedToUDTMember : IConvertToUDTMember
    //{
    //    private IEncapsulatableField _field;
    //    public ConvertedToUDTMember(IEncapsulatableField field)
    //    {

    //    }
    //}

    public class EncapsulateFieldCandidate : IEncapsulateFieldCandidate, IEncapsulateFieldCandidateValidations, IConvertToUDTMember//, IAssignNoConflictNames
    {
        protected Declaration _target;
        protected QualifiedModuleName _qmn;
        protected int _hashCode;
        private string _identifierName;
        protected EncapsulationIdentifiers _fieldAndProperty;

        public EncapsulateFieldCandidate(Declaration declaration, IValidateVBAIdentifiers identifierValidator) // Predicate<string> nameValidator/*, IValidateEncapsulateFieldNames validator*/)
        {
            _target = declaration;
            NameValidator = identifierValidator;

            _fieldAndProperty = new EncapsulationIdentifiers(declaration.IdentifierName, identifierValidator); // (string name) => nameValidator(name)); // (string name) => NamesValidator.IsValidVBAIdentifier(name, DeclarationType.Property, out _));
            IdentifierName = declaration.IdentifierName;
            FieldAsTypeName = declaration.AsTypeName;
            PropertyAsTypeName = declaration.AsTypeName;
            _qmn = declaration.QualifiedModuleName;
            NewPropertyAccessor = AccessorMember.Field;
            //_accessoryInProperty = _fieldAndProperty.Field;
            ReferenceAccessor = AccessorMember.Property;

            CanBeReadWrite = true;

            _hashCode = ($"{_qmn.Name}.{declaration.IdentifierName}").GetHashCode();

            ImplementLet = true;
            ImplementSet = false;
            if (_target.IsEnumField())
            {
                //5.3.1 The declared type of a function declaration may not be a private enum name.
                if (_target.AsTypeDeclaration.HasPrivateAccessibility())
                {
                    PropertyAsTypeName = Tokens.Long;
                }
            }
            else if (_target.AsTypeName.Equals(Tokens.Variant)
                && !_target.IsArray)
            {
                ImplementSet = true;
            }
            else if (Declaration.IsObject)
            {
                ImplementLet = false;
                ImplementSet = true;
            }
        }

        protected Dictionary<IdentifierReference, (ParserRuleContext, string)> IdentifierReplacements { get; } = new Dictionary<IdentifierReference, (ParserRuleContext, string)>();

        public Declaration Declaration => _target;

        public string AsTypeName => _target.AsTypeName;

        public virtual string FieldIdentifier
        {
            get => _fieldAndProperty.Field;
            set => _fieldAndProperty.Field = value;
        }
        public string FieldAsTypeName { set; get; }

        //private string _accessoryInProperty;
        //public virtual  string AccessorInProperty //{ get; }
        //    => _fieldAndProperty.Field; // AccessorMemberToContent(NewPropertyAccessor);
        //public string AccessorLocalReference { /*set;*/ get; }
        //public string AccessorExternalReference { set; get; }

        public virtual IValidateVBAIdentifiers NameValidator { set; get; }

        public virtual IEncapsulateFieldConflictFinder ConflictFinder { set; get; }

        public bool HasConflictingPropertyIdentifier
            => ConflictFinder.HasConflictingIdentifier(this, DeclarationType.Property, out var errorMessage);

        public bool HasConflictingFieldIdentifier
            => ConflictFinder.HasConflictingIdentifier(this, DeclarationType.Variable, out var errorMessage);

        public virtual bool TryValidateEncapsulationAttributes(out string errorMessage)
        {
            return ConflictFinder.TryValidateEncapsulationAttributes(this, out errorMessage);
        }

        public virtual IEnumerable<KeyValuePair<IdentifierReference, (ParserRuleContext, string)>> ReferenceReplacements
        {
            get
            {
                var results = new List<KeyValuePair<IdentifierReference, (ParserRuleContext, string)>>();
                foreach (var replacement in IdentifierReplacements)
                {
                    var kv = new KeyValuePair<IdentifierReference, (ParserRuleContext, string)>
                        (replacement.Key, replacement.Value);
                    results.Add(kv);
                }
                return results;
            }
        }

        protected virtual void SetReferenceRewriteContent(IdentifierReference idRef, string replacementText)
        {
            if (IdentifierReplacements.ContainsKey(idRef))
            {
                IdentifierReplacements[idRef] = (idRef.Context, replacementText);
                return;
            }
            IdentifierReplacements.Add(idRef, (idRef.Context, replacementText));
        }

        public virtual string TargetID => _target?.IdentifierName ?? IdentifierName;

        protected bool _encapsulateFlag;
        public virtual bool EncapsulateFlag
        {
            set
            {
                _encapsulateFlag = value;
                if (!_encapsulateFlag)
                {
                    PropertyIdentifier = _fieldAndProperty.DefaultPropertyName;
                    ConflictFinder.AssignNoConflictIdentifier(this, DeclarationType.Property);
                    ConflictFinder.AssignNoConflictIdentifier(this, DeclarationType.Variable);
                }
            }
            get => _encapsulateFlag;
        }

        public virtual bool IsReadOnly { set; get; }
        public bool CanBeReadWrite { set; get; }

        public virtual string PropertyName
        {
            get => PropertyIdentifier;
            set => PropertyIdentifier = value;
        }

        public override bool Equals(object obj)
        {
            return obj != null
                && obj is IEncapsulateFieldCandidate efc
                && $"{efc.QualifiedModuleName.Name}.{efc.TargetID}" == $"{_qmn.Name}.{IdentifierName}";
        }

        public override int GetHashCode() => _hashCode;
        /*
         * IConvertToUDTMemberInterface
         */
        public virtual string AccessorInProperty
            => $"{ObjectStateUDT.FieldIdentifier}.{UDTMemberIdentifier}";

        public virtual string AccessorLocalReference
            => $"{ObjectStateUDT.FieldIdentifier}.{PropertyIdentifier}";

        public string AccessorExternalReference { set; get; }
        public string PropertyIdentifier
        {
            get => _fieldAndProperty.Property;
            set
            {
                _fieldAndProperty.Property = value;
                UDTMemberIdentifier = value;

                TryRestoreNewFieldNameAsOriginalFieldIdentifierName();
            }
        }

        public string UDTMemberIdentifier { set; get; }

        public virtual string UDTMemberDeclaration
            => $"{PropertyIdentifier} {Tokens.As} {FieldAsTypeName}";

        public IObjectStateUDT ObjectStateUDT { set; get; }

        private void TryRestoreNewFieldNameAsOriginalFieldIdentifierName()
        {
            var canNowUseOriginalFieldName = !_fieldAndProperty.TargetFieldName.IsEquivalentVBAIdentifierTo(_fieldAndProperty.Property)
                && !ConflictFinder.IsConflictingProposedIdentifier(_fieldAndProperty.TargetFieldName, this, DeclarationType.Variable);

            if (canNowUseOriginalFieldName)
            {
                _fieldAndProperty.Field = _fieldAndProperty.TargetFieldName;
                return;
            }

            if (_fieldAndProperty.Field.IsEquivalentVBAIdentifierTo(_fieldAndProperty.TargetFieldName))
            {
                _fieldAndProperty.Field = _fieldAndProperty.DefaultNewFieldName;
                var isConflictingFieldIdentifier = ConflictFinder.HasConflictingIdentifier(this, DeclarationType.Variable, out _);
                for (var count = 1; count < 10 && isConflictingFieldIdentifier; count++)
                {
                    FieldIdentifier = FieldIdentifier.IncrementEncapsulationIdentifier();
                    isConflictingFieldIdentifier = ConflictFinder.HasConflictingIdentifier(this, DeclarationType.Variable, out _);
                }
            }
        }

        public string PropertyAsTypeName { get; set; }

        public QualifiedModuleName QualifiedModuleName => _qmn;

        public string IdentifierName
        {
            get => Declaration?.IdentifierName ?? _identifierName;
            set => _identifierName = value;
        }

        public virtual string ReferenceQualifier { set; get; }

        public string ParameterName => _fieldAndProperty.SetLetParameter;

        private bool _implLet;
        public bool ImplementLet { get => !IsReadOnly && _implLet; set => _implLet = value; }

        private bool _implSet;
        public bool ImplementSet { get => !IsReadOnly && _implSet; set => _implSet = value; }

        public bool ConvertFieldToUDTMember { set; get; }

        public EncapsulateFieldStrategy EncapsulateFieldStrategy { set; get; } = EncapsulateFieldStrategy.UseBackingFields;

        public virtual IEnumerable<IPropertyGeneratorAttributes> PropertyAttributeSets
            => new List<IPropertyGeneratorAttributes>() { AsPropertyAttributeSet };

        public virtual void LoadFieldReferenceContextReplacements(string referenceQualifier = null)
        {
            ReferenceQualifier = referenceQualifier;
            foreach (var idRef in Declaration.References)
            {
                var replacementText = RequiresAccessQualification(idRef)
                    ? $"{QualifiedModuleName.ComponentName}.{ReferenceForPreExistingReferences}"
                    : ReferenceForPreExistingReferences;

                SetReferenceRewriteContent(idRef, replacementText);
            }
        }

        protected AccessorMember NewPropertyAccessor { set; get; }

        protected AccessorMember ReferenceAccessor { set; get; }

        protected virtual string ReferenceWithinNewProperty 
            => AccessorMemberToContent(NewPropertyAccessor);

        protected virtual string ReferenceForPreExistingReferences 
            => AccessorMemberToContent(ReferenceAccessor);

        private string AccessorMemberToContent(AccessorMember accessorMember)
        {
            if ((ReferenceQualifier?.Length ?? 0) > 0)
            {
                return $"{ReferenceQualifier}.{PropertyIdentifier}";
            }

            return accessorMember == AccessorMember.Field
                ? FieldIdentifier
                : PropertyIdentifier;
        }

        protected virtual IPropertyGeneratorAttributes AsPropertyAttributeSet
        {
            get
            {
                return new PropertyAttributeSet()
                {
                    PropertyName = PropertyIdentifier,
                    BackingField = ReferenceWithinNewProperty,
                    AsTypeName = PropertyAsTypeName,
                    ParameterName = ParameterName,
                    GenerateLetter = ImplementLet,
                    GenerateSetter = ImplementSet,
                    UsesSetAssignment = Declaration.IsObject,
                    IsUDTProperty = false
                };
            }
        }

        protected virtual bool RequiresAccessQualification(IdentifierReference idRef)
        {
            var isLHSOfMemberAccess =
                        (idRef.Context.Parent is VBAParser.MemberAccessExprContext
                            || idRef.Context.Parent is VBAParser.WithMemberAccessExprContext)
                        && !(idRef.Context == idRef.Context.Parent.GetChild(0));

            return idRef.QualifiedModuleName != idRef.Declaration.QualifiedModuleName
                        && !isLHSOfMemberAccess;
        }
    }
}
