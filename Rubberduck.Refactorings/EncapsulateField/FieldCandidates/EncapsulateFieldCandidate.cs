using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.EncapsulateField.Extensions;
using Rubberduck.VBEditor;
using System.Collections.Generic;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldDeclaration
    {
        string IdentifierName { get; }
        QualifiedModuleName QualifiedModuleName { get; }
        string AsTypeName { get; }
        Selection Selection { get; }
        Accessibility Accessibility { get; }
    }

    public interface IEncapsulateFieldCandidate : IEncapsulateFieldDeclaration
    {
        Declaration Declaration { get; }
        //string IdentifierName { get; }
        string TargetID { get; }
        bool IsReadOnly { get; set; }
        bool EncapsulateFlag { get; set; }
        string FieldIdentifier { set; get; }
        bool CanBeReadWrite { set; get; }
        //QualifiedModuleName QualifiedModuleName { get; }
        string PropertyName { get; set; }
        string AsTypeName_Field { get; set; }
        string AsTypeName_Property { get; set; }
        string ParameterName { get; }
        bool ImplementLet { get; set; }
        bool ImplementSet { get; set; }
        IEnumerable<IPropertyGeneratorAttributes> PropertyAttributeSets { get; }
        string AsUDTMemberDeclaration { get; }
        IEnumerable<KeyValuePair<IdentifierReference, (ParserRuleContext, string)>> ReferenceReplacements { get; }
        void SetReferenceRewriteContent(IdentifierReference idRef, string replacementText);
        string ReferenceQualifier { set; get; }
        string ReferenceWithinNewProperty { get; }
        void StageFieldReferenceReplacements(IObjectStateUDT stateUDT = null);
        AccessorTokens PropertyAccessor { set; get; }
        AccessorTokens ReferenceAccessor { set; get; }
        bool TryValidateEncapsulationAttributes(out string errorMessage);

    }

    public enum AccessorTokens { Field, Property }

    public interface IEncapsulateFieldCandidateValidations
    {
        bool HasConflictingPropertyIdentifier { get; }
        bool HasConflictingFieldIdentifier { get; }
    }


    public class EncapsulateFieldCandidate : IEncapsulateFieldCandidate, IEncapsulateFieldCandidateValidations
    {
        protected Declaration _target;
        protected QualifiedModuleName _qmn;
        protected int _hashCode;
        private string _identifierName;
        protected IEncapsulateFieldNamesValidator _validator;
        protected EncapsulationIdentifiers _fieldAndProperty;

        public EncapsulateFieldCandidate(Declaration declaration, IEncapsulateFieldNamesValidator validator)
            : this(declaration.IdentifierName, declaration.AsTypeName, declaration.QualifiedModuleName, validator)
        {
            _target = declaration;
        }

        public EncapsulateFieldCandidate(string identifier, string asTypeName, QualifiedModuleName qmn, IEncapsulateFieldNamesValidator validator)
        {
            _target = null;

            _fieldAndProperty = new EncapsulationIdentifiers(identifier);
            IdentifierName = identifier;
            AsTypeName_Field = asTypeName;
            AsTypeName_Property = asTypeName;
            _qmn = qmn;
            PropertyAccessor = AccessorTokens.Field;
            ReferenceAccessor = AccessorTokens.Property;

            _validator = validator;

            ImplementLet = true;
            ImplementSet = false;
            CanBeReadWrite = true;

            _hashCode = ($"{_qmn.Name}.{identifier}").GetHashCode();
        }

        public virtual void StageFieldReferenceReplacements(IObjectStateUDT stateUDT = null)
        {
            PropertyAccessor = stateUDT is null ? AccessorTokens.Field : AccessorTokens.Property;
            ReferenceAccessor = AccessorTokens.Property;
            ReferenceQualifier = stateUDT?.FieldIdentifier ?? null;
            LoadFieldReferenceContextReplacements();
        }

        protected Dictionary<IdentifierReference, (ParserRuleContext, string)> IdentifierReplacements { get; } = new Dictionary<IdentifierReference, (ParserRuleContext, string)>();

        public Declaration Declaration => _target;

        public Selection Selection => _target.Selection;

        public Accessibility Accessibility => _target.Accessibility;

        public string AsTypeName => _target.AsTypeName;

        public bool HasConflictingPropertyIdentifier 
            => _validator.HasConflictingIdentifier(this, DeclarationType.Property, out var errorMessage);

        public bool HasConflictingFieldIdentifier
            => _validator.HasConflictingIdentifier(this, DeclarationType.Variable, out var errorMessage);

        public virtual bool TryValidateEncapsulationAttributes(out string errorMessage)
        {
            errorMessage = string.Empty;
            if (!EncapsulateFlag) { return true; }

            if (!_validator.IsValidVBAIdentifier(PropertyName, DeclarationType.Property, out errorMessage))
            {
                return false;
            }

            if (!_validator.IsSelfConsistent(this, out errorMessage))
            {
                return false;
            }

            if (_validator.HasConflictingIdentifier(this, DeclarationType.Property, out errorMessage))
            {
                return false;
            }

            if (_validator.HasConflictingIdentifier(this, DeclarationType.Variable, out errorMessage))
            {
                return false;
            }
            return true;
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

        public virtual void SetReferenceRewriteContent(IdentifierReference idRef, string replacementText)
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
                    PropertyName = _fieldAndProperty.DefaultPropertyName;
                    _validator.AssignNoConflictIdentifier(this, DeclarationType.Property);
                    _validator.AssignNoConflictIdentifier(this, DeclarationType.Variable);
                }
            }
            get => _encapsulateFlag;
        }

        public virtual bool IsReadOnly { set; get; }
        public bool CanBeReadWrite { set; get; }

        public virtual string FieldIdentifier
        {
            get => _fieldAndProperty.Field;
            set => _fieldAndProperty.Field = value;
        }

        public virtual string PropertyName
        {
            get => _fieldAndProperty.Property;
            set
            {
                _fieldAndProperty.Property = value;

                TryRestoreNewFieldNameAsOriginalFieldIdentifierName();
            }
        }

        public override bool Equals(object obj)
        {
            return obj != null 
                && obj is IEncapsulateFieldCandidate 
                && obj.GetHashCode() == GetHashCode();
        }

        public override int GetHashCode() => _hashCode;

        //The preferred NewFieldName is the original Identifier
        private void TryRestoreNewFieldNameAsOriginalFieldIdentifierName()
        {
            var canNowUseOriginalFieldName = !_fieldAndProperty.TargetFieldName.IsEquivalentVBAIdentifierTo(_fieldAndProperty.Property)
                && !_validator.IsConflictingFieldIdentifier(_fieldAndProperty.TargetFieldName, this, DeclarationType.Variable);

            if (canNowUseOriginalFieldName)
            {
                _fieldAndProperty.Field = _fieldAndProperty.TargetFieldName;
                return;
            }

            if (_fieldAndProperty.Field.IsEquivalentVBAIdentifierTo(_fieldAndProperty.TargetFieldName))
            {
                _fieldAndProperty.Field = _fieldAndProperty.DefaultNewFieldName;
                var isConflictingFieldIdentifier = _validator.HasConflictingIdentifier(this, DeclarationType.Variable, out _);
                for (var count = 1; count < 10 && isConflictingFieldIdentifier; count++)
                {
                    FieldIdentifier = FieldIdentifier.IncrementEncapsulationIdentifier();
                    isConflictingFieldIdentifier = _validator.HasConflictingIdentifier(this, DeclarationType.Variable, out _);
                }
            }
        }

        public string AsTypeName_Field { set; get; }

        public string AsTypeName_Property { get; set; }

        public QualifiedModuleName QualifiedModuleName => _qmn;

        public string IdentifierName
        {
            get => Declaration?.IdentifierName ?? _identifierName;
            set => _identifierName = value;
        }

        public string ParameterName => _fieldAndProperty.SetLetParameter;

        private bool _implLet;
        public bool ImplementLet { get => !IsReadOnly && _implLet; set => _implLet = value; }

        private bool _implSet;
        public bool ImplementSet { get => !IsReadOnly && _implSet; set => _implSet = value; }

        public AccessorTokens PropertyAccessor { set; get; }

        public AccessorTokens ReferenceAccessor { set; get; }

        protected string _referenceQualifier;
        public virtual string ReferenceQualifier
        {
            set => _referenceQualifier = value;
            get => _referenceQualifier;
        }

        public virtual string ReferenceWithinNewProperty => AccessorTokenToContent(PropertyAccessor);

        protected virtual string ReferenceForPreExistingReferences => AccessorTokenToContent(ReferenceAccessor);

        private string AccessorTokenToContent(AccessorTokens token)
        {
            var accessor = token == AccessorTokens.Field
                ? FieldIdentifier
                : PropertyName;

            if ((ReferenceQualifier?.Length ?? 0) > 0)
            {
                return $"{ReferenceQualifier}.{accessor}";
            }
            return accessor;
        }

        public virtual string AsUDTMemberDeclaration 
            => $"{PropertyName} {Tokens.As} {AsTypeName_Field}";

        public virtual IEnumerable<IPropertyGeneratorAttributes> PropertyAttributeSets 
            => new List<IPropertyGeneratorAttributes>() { AsPropertyAttributeSet };

        protected virtual IPropertyGeneratorAttributes AsPropertyAttributeSet
        {
            get
            {
                return new PropertyAttributeSet()
                {
                    PropertyName = PropertyName,
                    BackingField = ReferenceWithinNewProperty,
                    AsTypeName = AsTypeName_Property,
                    ParameterName = ParameterName,
                    GenerateLetter = ImplementLet,
                    GenerateSetter = ImplementSet,
                    UsesSetAssignment = Declaration.IsObject
                };
            }
        }

        protected virtual void LoadFieldReferenceContextReplacements()
        {
            var field = this;
            foreach (var idRef in field.Declaration.References)
            {
                var replacementText = RequiresAccessQualification(idRef)
                    ? $"{field.QualifiedModuleName.ComponentName}.{field.ReferenceForPreExistingReferences}"
                    : field.ReferenceForPreExistingReferences;

                field.SetReferenceRewriteContent(idRef, replacementText);
            }
        }

        protected bool RequiresAccessQualification(IdentifierReference idRef)
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
