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
        string NewFieldName { set; get; }
        bool CanBeReadWrite { set; get; }
        bool IsUDTMember { get; }
        bool HasValidEncapsulationAttributes { get; }
        QualifiedModuleName QualifiedModuleName { get; }
        IEnumerable<IdentifierReference> References { get; }
        string PropertyName { get; set; }
        string AsTypeName { get; set; }
        string ParameterName { get; }
        bool ImplementLetSetterType { get; set; }
        bool ImplementSetSetterType { get; set; }
        Func<string> PropertyAccessExpression { set; get; }
        bool FieldNameIsExemptFromValidation { get; }
        Func<string> ReferenceExpression { set; get; }
        IEnumerable<IPropertyGeneratorSpecification> PropertyGenerationSpecs { get; }
        IEnumerable<KeyValuePair<IdentifierReference, RewriteReplacePair>> ReferenceReplacements { get; }
        void LoadReferenceExpressionChanges();
        void SetReferenceRewriteContent(IdentifierReference idRef, string replacementText);
        RewriteReplacePair ReferenceReplacement(IdentifierReference idRef);
    }

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
            ImplementLetSetterType = true;
            ImplementSetSetterType = false;
            CanBeReadWrite = true;
        }

        public EncapsulateFieldCandidate(string identifier, string asTypeName, QualifiedModuleName qmn, IEncapsulateFieldNamesValidator validator)
        {
            _target = null;

            _fieldAndProperty = new EncapsulationIdentifiers(identifier);
            IdentifierName = identifier;
            AsTypeName = asTypeName;
            _qmn = qmn;
            PropertyAccessExpression = () => NewFieldName;
            ReferenceExpression = () => PropertyName;

            _validator = validator;
        }

        protected Dictionary<IdentifierReference, RewriteReplacePair> IdentifierReplacements { get; } = new Dictionary<IdentifierReference, RewriteReplacePair>();

        public Declaration Declaration => _target;


        public bool HasVBACompliantPropertyIdentifier => _validator.IsValidVBAIdentifier(PropertyName, DeclarationType.Property);

        public bool HasVBACompliantFieldIdentifier => _validator.IsValidVBAIdentifier(NewFieldName, Declaration?.DeclarationType ?? DeclarationType.Variable);

        public bool HasVBACompliantParameterIdentifier => _validator.IsValidVBAIdentifier(NewFieldName, Declaration?.DeclarationType ?? DeclarationType.Variable);

        public virtual bool IsSelfConsistent => _validator.IsValidVBAIdentifier(PropertyName, DeclarationType.Property)
                            && !(PropertyName.EqualsVBAIdentifier(NewFieldName)
                                    || PropertyName.EqualsVBAIdentifier(ParameterName)
                                    || NewFieldName.EqualsVBAIdentifier(ParameterName));

        public bool HasConflictingPropertyIdentifier 
            => _validator.HasConflictingPropertyIdentifier(this);

        public bool HasConflictingFieldIdentifier
            => _validator.HasConflictingFieldIdentifier(this);

        public bool HasValidEncapsulationAttributes
        {
            get
            {
                if (!EncapsulateFlag) { return true; }

                return IsSelfConsistent
                    && !_validator.HasConflictingPropertyIdentifier(this);
            }
        }

        public virtual IEnumerable<KeyValuePair<IdentifierReference, RewriteReplacePair>> ReferenceReplacements
        {
            get
            {
                var results = new List<KeyValuePair<IdentifierReference, RewriteReplacePair>>();
                foreach (var replacement in IdentifierReplacements)
                {
                    var kv = new KeyValuePair<IdentifierReference, RewriteReplacePair>
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
                IdentifierReplacements[idRef] = new RewriteReplacePair(replacementText, idRef.Context);
                return;
            }
            IdentifierReplacements.Add(idRef, new RewriteReplacePair(replacementText, idRef.Context));
        }

        public RewriteReplacePair ReferenceReplacement(IdentifierReference idRef)
        {
            return IdentifierReplacements.Single(r => r.Key == idRef).Value;
        }

        public virtual string TargetID => _target?.IdentifierName ?? IdentifierName;

        public bool EncapsulateFlag { set; get; }
        public bool IsReadOnly { set; get; }
        public bool CanBeReadWrite { set; get; }

        public virtual string NewFieldName
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
            var canNowUseOriginalFieldName = !_validator.IsConflictingFieldIdentifier(_fieldAndProperty.TargetFieldName, this);

            if (canNowUseOriginalFieldName)
            {
                _fieldAndProperty.Field = _fieldAndProperty.TargetFieldName;
                return;
            }

            if (_fieldAndProperty.Field.EqualsVBAIdentifier(_fieldAndProperty.TargetFieldName))
            {
                _fieldAndProperty.Field = _fieldAndProperty.DefaultNewFieldName;
                var isConflictingFieldIdentifier = _validator.HasConflictingFieldIdentifier(this);
                for (var count = 1; count < 10 && isConflictingFieldIdentifier; count++)
                {
                    NewFieldName = EncapsulationIdentifiers.IncrementIdentifier(NewFieldName);
                    isConflictingFieldIdentifier = _validator.HasConflictingFieldIdentifier(this);
                }
            }
        }

        public string AsTypeName { set; get; }

        public bool IsUDTMember => _target?.DeclarationType.Equals(DeclarationType.UserDefinedTypeMember) ?? false;

        public QualifiedModuleName QualifiedModuleName => _qmn;

        public virtual IEnumerable<IdentifierReference> References => Declaration?.References ?? Enumerable.Empty<IdentifierReference>();

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


        public Func<string> PropertyAccessExpression { set; get; }

        public virtual IEnumerable<IPropertyGeneratorSpecification> PropertyGenerationSpecs 
            => new List<IPropertyGeneratorSpecification>() { AsPropertyGeneratorSpec };

        protected virtual IPropertyGeneratorSpecification AsPropertyGeneratorSpec
        {
            get
            {
                return new PropertyGeneratorSpecification()
                {
                    PropertyName = PropertyName,
                    BackingField = PropertyAccessExpression(),
                    AsTypeName = AsTypeName,
                    ParameterName = ParameterName,
                    GenerateLetter = ImplementLetSetterType,
                    GenerateSetter = ImplementSetSetterType,
                    UsesSetAssignment = Declaration.IsObject
                };
            }
        }

        public virtual void LoadReferenceExpressionChanges()
        {
            LoadFieldReferenceExpressions();
        }

        protected virtual void LoadFieldReferenceExpressions()
        {
            var field = this;
            foreach (var idRef in field.References)
            {
                var replacementText = RequiresAccessQualification(idRef)
                    ? $"{field.QualifiedModuleName.ComponentName}.{field.ReferenceExpression()}"
                    : field.ReferenceExpression();

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

        public Func<string> ReferenceExpression { set; get; }

        public bool FieldNameIsExemptFromValidation 
            => Declaration?.DeclarationType.Equals(DeclarationType.UserDefinedTypeMember) ?? false;
    }
}
