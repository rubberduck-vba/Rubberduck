using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Common;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldCandidate
    {
        Declaration Declaration { get; }
        string IdentifierName { get; }
        string TargetID { get; }
        bool IsReadOnly { get; set; }
        bool EncapsulateFlag { get; set; }
        string FieldIdentifier { set; get; }
        bool CanBeReadWrite { set; get; }
        bool HasValidEncapsulationAttributes { get; }
        QualifiedModuleName QualifiedModuleName { get; }
        string PropertyName { get; set; }
        string AsTypeName { get; set; }
        string ParameterName { get; }
        bool ImplementLetSetterType { get; set; }
        bool ImplementSetSetterType { get; set; }
        IEnumerable<IPropertyGeneratorAttributes> PropertyAttributeSets { get; }
        IEnumerable<KeyValuePair<IdentifierReference, (ParserRuleContext, string)>> ReferenceReplacements { get; }
        void SetReferenceRewriteContent(IdentifierReference idRef, string replacementText);
        string ReferenceQualifier { set; get; }
        string ReferenceWithinNewProperty { get; }
        void StageFieldReferenceReplacements(IStateUDT stateUDT = null);
    }

    public enum AccessorTokens { Field, Property }

    public interface IEncapsulateFieldCandidateValidations
    {
        bool HasVBACompliantPropertyIdentifier { get; }
        bool HasVBACompliantFieldIdentifier { get; }
        bool HasVBACompliantParameterIdentifier { get; }
        bool IsSelfConsistent { get; }
        bool HasConflictingPropertyIdentifier { get; }
        bool HasConflictingFieldIdentifier { get; }
    }


    public class EncapsulateFieldCandidate : IEncapsulateFieldCandidate, IEncapsulateFieldCandidateValidations
    {
        protected Declaration _target;
        protected QualifiedModuleName _qmn;
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
            AsTypeName = asTypeName;
            _qmn = qmn;
            PropertyAccessor = AccessorTokens.Field;
            ReferenceAccessor = AccessorTokens.Property;

            _validator = validator;

            ImplementLetSetterType = true;
            ImplementSetSetterType = false;
            CanBeReadWrite = true;
        }

        public virtual void StageFieldReferenceReplacements(IStateUDT stateUDT = null)
        {
            PropertyAccessor = stateUDT is null ? AccessorTokens.Field : AccessorTokens.Property;
            ReferenceAccessor = AccessorTokens.Property;
            ReferenceQualifier = stateUDT?.FieldIdentifier ?? null;
            LoadFieldReferenceContextReplacements();
        }

        protected Dictionary<IdentifierReference, (ParserRuleContext, string)> IdentifierReplacements { get; } = new Dictionary<IdentifierReference, (ParserRuleContext, string)>();

        public Declaration Declaration => _target;

        public bool HasVBACompliantPropertyIdentifier => _validator.IsValidVBAIdentifier(PropertyName, DeclarationType.Property);

        public bool HasVBACompliantFieldIdentifier => _validator.IsValidVBAIdentifier(FieldIdentifier, Declaration?.DeclarationType ?? DeclarationType.Variable);

        public bool HasVBACompliantParameterIdentifier => _validator.IsValidVBAIdentifier(FieldIdentifier, Declaration?.DeclarationType ?? DeclarationType.Variable);

        public virtual bool IsSelfConsistent => _validator.IsValidVBAIdentifier(PropertyName, DeclarationType.Property)
                            && !(PropertyName.EqualsVBAIdentifier(FieldIdentifier)
                                    || PropertyName.EqualsVBAIdentifier(ParameterName)
                                    || FieldIdentifier.EqualsVBAIdentifier(ParameterName));

        public bool HasConflictingPropertyIdentifier 
            => _validator.HasConflictingIdentifier(this, DeclarationType.Property);

        public bool HasConflictingFieldIdentifier
            => _validator.HasConflictingIdentifier(this, DeclarationType.Variable);

        public bool HasValidEncapsulationAttributes
        {
            get
            {
                if (!EncapsulateFlag) { return true; }

                return IsSelfConsistent
                    && !_validator.HasConflictingIdentifier(this, DeclarationType.Property);
            }
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

        public bool EncapsulateFlag { set; get; }
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

        //The preferred NewFieldName is the original Identifier
        private void TryRestoreNewFieldNameAsOriginalFieldIdentifierName()
        {
            var canNowUseOriginalFieldName = !_fieldAndProperty.TargetFieldName.EqualsVBAIdentifier(_fieldAndProperty.Property)
                && !_validator.IsConflictingFieldIdentifier(_fieldAndProperty.TargetFieldName, this, DeclarationType.Variable);

            if (canNowUseOriginalFieldName)
            {
                _fieldAndProperty.Field = _fieldAndProperty.TargetFieldName;
                return;
            }

            if (_fieldAndProperty.Field.EqualsVBAIdentifier(_fieldAndProperty.TargetFieldName))
            {
                _fieldAndProperty.Field = _fieldAndProperty.DefaultNewFieldName;
                var isConflictingFieldIdentifier = _validator.HasConflictingIdentifier(this, DeclarationType.Variable);
                for (var count = 1; count < 10 && isConflictingFieldIdentifier; count++)
                {
                    FieldIdentifier = FieldIdentifier.IncrementIdentifier();
                    isConflictingFieldIdentifier = _validator.HasConflictingIdentifier(this, DeclarationType.Variable);
                }
            }
        }

        public string AsTypeName { set; get; }

        public QualifiedModuleName QualifiedModuleName => _qmn;

        public string IdentifierName
        {
            get => Declaration?.IdentifierName ?? _identifierName;
            set => _identifierName = value;
        }

        public string ParameterName => _fieldAndProperty.SetLetParameter;

        private bool _implLet;
        public bool ImplementLetSetterType { get => !IsReadOnly && _implLet; set => _implLet = value; }

        private bool _implSet;
        public bool ImplementSetSetterType { get => !IsReadOnly && _implSet; set => _implSet = value; }


        protected AccessorTokens PropertyAccessor { set; get; }

        protected AccessorTokens ReferenceAccessor { set; get; }

        protected string _referenceQualifier;
        public virtual string ReferenceQualifier
        {
            set => _referenceQualifier = value;
            get => _referenceQualifier;
        }

        public virtual string ReferenceWithinNewProperty => AccessorTokenToContent(PropertyAccessor);

        protected virtual string ReferenceForPreExistingReferences => AccessorTokenToContent(ReferenceAccessor);

        protected string AccessorTokenToContent(AccessorTokens token)
        {
            var accessor = string.Empty;
            switch (token)
            {
                case AccessorTokens.Field:
                    accessor = FieldIdentifier;
                    break;
                case AccessorTokens.Property:
                    accessor = PropertyName;
                    break;
                default:
                    throw new ArgumentException();
            }

            if ((ReferenceQualifier?.Length ?? 0) > 0)
            {
                return $"{ReferenceQualifier}.{accessor}";
            }
            return accessor;
        }

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
                    AsTypeName = AsTypeName,
                    ParameterName = ParameterName,
                    GenerateLetter = ImplementLetSetterType,
                    GenerateSetter = ImplementSetSetterType,
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
