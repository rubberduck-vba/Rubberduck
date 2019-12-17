using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IUserDefinedTypeMemberCandidate : IEncapsulateFieldCandidate
    {
        IUserDefinedTypeCandidate Parent { get; }
        bool IncludeParentNameWithPropertyIdentifier { set; get; }
        IPropertyGeneratorAttributes AsPropertyGeneratorSpec { get; }
        Dictionary<IdentifierReference, (ParserRuleContext, string)> IdentifierReplacements { get; }
        IEnumerable<IdentifierReference> FieldRelatedReferences(IUserDefinedTypeCandidate field);
    }

    //public class UserDefinedTypeMemberCandidate : EncapsulateFieldCandidate, IUserDefinedTypeMemberCandidate
    public class UserDefinedTypeMemberCandidate : /*IEncapsulateFieldCandidate,*/ IUserDefinedTypeMemberCandidate
    {
        private readonly IEncapsulateFieldNamesValidator _validator;
        public UserDefinedTypeMemberCandidate(IEncapsulateFieldCandidate efd, IUserDefinedTypeCandidate udtVariable, IEncapsulateFieldNamesValidator validator)
        {
            _decoratedField = efd;
            Parent = udtVariable;
            _validator = validator;
            PropertyName = IdentifierName;
            PropertyAccessor = AccessorTokens.Property;
            ReferenceAccessor = AccessorTokens.Property;
        }
        //public UserDefinedTypeMemberCandidate(Declaration target, IUserDefinedTypeCandidate udtVariable, IEncapsulateFieldNamesValidator validator)
        //    : base(target, validator)
        //{
        //    Parent = udtVariable;

        //    PropertyName = IdentifierName;
        //    PropertyAccessor = AccessorTokens.Property;
        //    ReferenceAccessor = AccessorTokens.Property;
        //}

        private IEncapsulateFieldCandidate _decoratedField;

        public IUserDefinedTypeCandidate Parent { private set; get; }

        private bool _includeParentNameWithPropertyIdentifier;
        public bool IncludeParentNameWithPropertyIdentifier
        {
            get => _includeParentNameWithPropertyIdentifier;
            set
            {
                _includeParentNameWithPropertyIdentifier = value;
                PropertyName = _includeParentNameWithPropertyIdentifier
                    ? $"{Parent.IdentifierName.Capitalize()}_{IdentifierName}"
                    : IdentifierName;
            }
        }

        public void StageFieldReferenceReplacements(IStateUDT stateUDT = null) { }

        private string _referenceQualifier;
        public string ReferenceQualifier
        {
            set => _referenceQualifier = value;
            get => Parent.ReferenceWithinNewProperty;
        }

        public string TargetID => $"{Parent.IdentifierName}.{IdentifierName}";

        public IEnumerable<IdentifierReference> FieldRelatedReferences(IUserDefinedTypeCandidate field)
            => GetUDTMemberReferencesForField(this, field);

        public void SetReferenceRewriteContent(IdentifierReference idRef, string replacementText)
        {
            Debug.Assert(idRef.Context.Parent is ParserRuleContext, "idRef.Context.Parent is not convertable to ParserRuleContext");

            if (IdentifierReplacements.ContainsKey(idRef))
            {
                IdentifierReplacements[idRef] = (idRef.Context.Parent as ParserRuleContext, replacementText);
                return;
            }
            IdentifierReplacements.Add(idRef, (idRef.Context.Parent as ParserRuleContext, replacementText));
        }

        public IPropertyGeneratorAttributes AsPropertyGeneratorSpec
        {
            get
            {
                if (_includeParentNameWithPropertyIdentifier)
                {
                    PropertyAccessor = AccessorTokens.Field;
                }

                return new PropertyAttributeSet()
                {
                    PropertyName = PropertyName,
                    BackingField = ReferenceWithinNewProperty,
                    AsTypeName = AsTypeName_Property,
                    ParameterName = ParameterName,
                    GenerateLetter = ImplementLetSetterType,
                    GenerateSetter = ImplementSetSetterType,
                    UsesSetAssignment = Declaration.IsObject
                };
            }
        }

        //public new Dictionary<IdentifierReference, (ParserRuleContext, string)> IdentifierReplacements { get; } = new Dictionary<IdentifierReference, (ParserRuleContext, string)>();
        public Dictionary<IdentifierReference, (ParserRuleContext, string)> IdentifierReplacements { get; } = new Dictionary<IdentifierReference, (ParserRuleContext, string)>();

        private IEnumerable<IdentifierReference> GetUDTMemberReferencesForField(IEncapsulateFieldCandidate udtMember, IUserDefinedTypeCandidate field)
        {
            var refs = new List<IdentifierReference>();
            foreach (var idRef in udtMember.Declaration.References)
            {
                if (idRef.Context.TryGetAncestor<VBAParser.MemberAccessExprContext>(out var mac))
                {
                    var LHS = mac.children.First();
                    switch (LHS)
                    {
                        case VBAParser.SimpleNameExprContext snec:
                            if (snec.GetText().Equals(field.IdentifierName))
                            {
                                refs.Add(idRef);
                            }
                            break;
                        case VBAParser.MemberAccessExprContext submac:
                            if (submac.children.Last() is VBAParser.UnrestrictedIdentifierContext ur && ur.GetText().Equals(field.IdentifierName))
                            {
                                refs.Add(idRef);
                            }
                            break;
                        case VBAParser.WithMemberAccessExprContext wmac:
                            if (wmac.children.Last().GetText().Equals(field.IdentifierName))
                            {
                                refs.Add(idRef);
                            }
                            break;
                    }
                }
                else if (idRef.Context.TryGetAncestor<VBAParser.WithMemberAccessExprContext>(out var wmac))
                {
                    var wm = wmac.GetAncestor<VBAParser.WithStmtContext>();
                    var Lexpr = wm.GetChild<VBAParser.LExprContext>();
                    if (Lexpr.GetText().Equals(field.IdentifierName))
                    {
                        refs.Add(idRef);
                    }
                }
            }
            return refs;
        }

        public Declaration Declaration => _decoratedField.Declaration;
        public string IdentifierName => _decoratedField.IdentifierName;
        //public string TargetID { get; }
        public bool IsReadOnly
        {
            set => _decoratedField.IsReadOnly = value;
            get => _decoratedField.IsReadOnly;
        }
        public bool EncapsulateFlag //{ get; set; }
        {
            set => _decoratedField.EncapsulateFlag = value;
            get => _decoratedField.EncapsulateFlag;
        }
        public string FieldIdentifier //{ set; get; }
        {
            set => _decoratedField.FieldIdentifier = value;
            get => _decoratedField.FieldIdentifier;
        }
        public bool CanBeReadWrite //{ set; get; }
        {
            set => _decoratedField.CanBeReadWrite = value;
            get => _decoratedField.CanBeReadWrite;
        }
        public bool HasValidEncapsulationAttributes 
            => _decoratedField.HasValidEncapsulationAttributes; // { get; }

        public QualifiedModuleName QualifiedModuleName 
            => _decoratedField.QualifiedModuleName; // { get; }

        public string PropertyName //{ get; set; }
        {
            set => _decoratedField.PropertyName = value;
            get => _decoratedField.PropertyName;
        }
        public string AsTypeName //{ get; set; }
        {
            set => _decoratedField.AsTypeName = value;
            get => _decoratedField.AsTypeName;
        }
        public string AsTypeName_Property //{ get; set; }
        {
            set => _decoratedField.AsTypeName_Property = value;
            get => _decoratedField.AsTypeName_Property;
        }
        public string ParameterName => _decoratedField.ParameterName; // { get; }
        public bool ImplementLetSetterType //{ get; set; }
        {
            set => _decoratedField.ImplementLetSetterType = value;
            get => _decoratedField.ImplementLetSetterType;
        }
        public bool ImplementSetSetterType //{ get; set; }
        {
            set => _decoratedField.ImplementSetSetterType = value;
            get => _decoratedField.ImplementSetSetterType;
        }
        public IEnumerable<IPropertyGeneratorAttributes> PropertyAttributeSets => _decoratedField.PropertyAttributeSets; // { get; }
        public string AsUDTMemberDeclaration { get; }
        public IEnumerable<KeyValuePair<IdentifierReference, (ParserRuleContext, string)>> ReferenceReplacements => _decoratedField.ReferenceReplacements; // { get; }
        //public void SetReferenceRewriteContent(IdentifierReference idRef, string replacementText);
        //public string ReferenceQualifier { set; get; }
        public string ReferenceWithinNewProperty => $"{Parent.ReferenceWithinNewProperty}.{_decoratedField.ReferenceWithinNewProperty}"; // { get; }
        //public string ReferenceWithinNewProperty => _field.ReferenceWithinNewProperty; // { get; }
        //public void StageFieldReferenceReplacements(IStateUDT stateUDT = null);
        public AccessorTokens PropertyAccessor //{ set; get; }
        {
            set => _decoratedField.PropertyAccessor = value;
            get => _decoratedField.PropertyAccessor;
        }

        public AccessorTokens ReferenceAccessor //{ set; get; }
        {
            set => _decoratedField.ReferenceAccessor = value;
            get => _decoratedField.ReferenceAccessor;
        }

        public string AccessorTokenToContent(AccessorTokens token) => _decoratedField.AccessorTokenToContent(token);
    }
}
