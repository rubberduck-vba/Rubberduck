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
        //IPropertyGeneratorSpecification AsPropertyGeneratorSpec { get; }
        IEnumerable<IPropertyGeneratorSpecification> PropertyGenerationSpecs { get; }
        IEnumerable<KeyValuePair<IdentifierReference, RewriteReplacePair>> ReferenceReplacements { get; }
        void LoadReferenceExpressionChanges();
        void AddReferenceReplacement(IdentifierReference idRef, string replacementText);
        RewriteReplacePair ReferenceReplacement(IdentifierReference idRef);
        RewriteReplacePair? FindRewriteReplacePair(IdentifierReference idRef);
    }

    public class EncapsulateFieldCandidate : IEncapsulateFieldCandidate
    {
        protected Declaration _target;
        protected QualifiedModuleName _qmn;
        private string _identifierName;
        private IEncapsulateFieldNamesValidator _validator;
        private Dictionary<IdentifierReference, string> _idRefRenames { set; get; } = new Dictionary<IdentifierReference, string>();
        private EncapsulationIdentifiers _fieldAndProperty;

        public EncapsulateFieldCandidate(Declaration declaration, IEncapsulateFieldNamesValidator validator)
        {
            _target = declaration;
            AsTypeName = _target.AsTypeName;

            _fieldAndProperty = new EncapsulationIdentifiers(_target);
            IdentifierName = _target.IdentifierName;
            _qmn = _target.QualifiedModuleName;
            PropertyAccessExpression = () => NewFieldName;
            ReferenceExpression = () => PropertyName;

            _validator = validator;
            _validator.ForceNonConflictEncapsulationAttributes(this, _target.QualifiedModuleName, _target);
        }

        public EncapsulateFieldCandidate(string identifier, string asTypeName, QualifiedModuleName qmn,/*IFieldEncapsulationAttributes attributes,*/ IEncapsulateFieldNamesValidator validator, bool neverEncapsulate = false)
        {
            _target = null;

            _fieldAndProperty = new EncapsulationIdentifiers(identifier, neverEncapsulate);
            IdentifierName = identifier;
            AsTypeName = asTypeName;
            _qmn = qmn;
            PropertyAccessExpression = () => NewFieldName;
            ReferenceExpression = () => PropertyName;

            _validator = validator;
            _validator.ForceNonConflictEncapsulationAttributes(this, qmn, _target);
        }

        protected Dictionary<IdentifierReference, RewriteReplacePair> IdentifierReplacements { get; } = new Dictionary<IdentifierReference, RewriteReplacePair>();

        public Declaration Declaration => _target;

        public bool HasValidEncapsulationAttributes
        {
            get
            {
                var declarationsToIgnore = _target != null ? new Declaration[] { _target } : Enumerable.Empty<Declaration>();
                var declarationType = _target != null ? _target.DeclarationType : DeclarationType.Variable;
                return _validator.HasValidEncapsulationAttributes(this, QualifiedModuleName, declarationsToIgnore, declarationType);
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


        public RewriteReplacePair? FindRewriteReplacePair(IdentifierReference idRef)
        {
            if (IdentifierReplacements.ContainsKey(idRef))
            {
                return IdentifierReplacements[idRef];
            }
            return null;
        }

        public virtual void AddReferenceReplacement(IdentifierReference idRef, string replacementText)
        {
            IdentifierReplacements.Add(idRef, new RewriteReplacePair(replacementText, idRef.Context));
        }

        public RewriteReplacePair ReferenceReplacement(IdentifierReference idRef) //, ParserRuleContextExtensions context)]
        {
            return IdentifierReplacements.Single(r => r.Key == idRef).Value;
        }

        public virtual string TargetID => _target?.IdentifierName ?? IdentifierName;

        public bool EncapsulateFlag { set; get; }
        public bool IsReadOnly { set; get; }
        public bool CanBeReadWrite { set; get; }

        public string PropertyName
        {
            get => _fieldAndProperty.Property;
            set => _fieldAndProperty.Property = value;
        }

        public bool IsEditableReadWriteFieldIdentifier { set; get; } = true;

        public string NewFieldName
        {
            get => _fieldAndProperty.Field;
            set => _fieldAndProperty.Field = value;
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

                field.AddReferenceReplacement(idRef, replacementText);
            }
        }

        protected bool RequiresAccessQualification(IdentifierReference idRef)
        {
            var isLHSOfMemberAccess =
                        (idRef.Context.Parent is VBAParser.MemberAccessExprContext
                            || idRef.Context.Parent is VBAParser.WithMemberAccessExprContext)
                        && !(idRef.Context == idRef.Context.Parent.GetChild(0));// is VBAParser.SimpleNameExprContext))

            return idRef.QualifiedModuleName != idRef.Declaration.QualifiedModuleName
                        && !isLHSOfMemberAccess;
        }

        public Func<string> ReferenceExpression { set; get; }

        public bool FieldNameIsExemptFromValidation 
            => Declaration?.DeclarationType.Equals(DeclarationType.UserDefinedTypeMember) ?? false;
    }
}
