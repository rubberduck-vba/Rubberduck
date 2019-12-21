using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IUserDefinedTypeMemberCandidate : IEncapsulateFieldCandidate
    {
        IUserDefinedTypeCandidate Parent { get; }
        IPropertyGeneratorAttributes AsPropertyGeneratorSpec { get; }
        Dictionary<IdentifierReference, (ParserRuleContext, string)> IdentifierReplacements { get; }
        IEnumerable<IdentifierReference> FieldRelatedReferences(IUserDefinedTypeCandidate field);
    }

    public class UserDefinedTypeMemberCandidate : IUserDefinedTypeMemberCandidate
    {
        private readonly IEncapsulateFieldNamesValidator _validator;
        private int _hashCode;
        public UserDefinedTypeMemberCandidate(IEncapsulateFieldCandidate efd, IUserDefinedTypeCandidate udtVariable, IEncapsulateFieldNamesValidator validator)
        {
            _decoratedField = efd;
            Parent = udtVariable;
            _validator = validator;
            PropertyName = IdentifierName;
            PropertyAccessor = AccessorTokens.Property;
            ReferenceAccessor = AccessorTokens.Property;
            _hashCode = ($"{efd.QualifiedModuleName.Name}.{efd.IdentifierName}").GetHashCode();
        }

        private IEncapsulateFieldCandidate _decoratedField;

        public IUserDefinedTypeCandidate Parent { private set; get; }

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

        public Dictionary<IdentifierReference, (ParserRuleContext, string)> IdentifierReplacements { get; } = new Dictionary<IdentifierReference, (ParserRuleContext, string)>();


        public override bool Equals(object obj)
        {
            return obj != null
                && obj is IUserDefinedTypeMemberCandidate
                && obj.GetHashCode() == GetHashCode();
        }

        public override int GetHashCode() => _hashCode;

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

        public bool TryValidateEncapsulationAttributes(out string errorMessage)
        {
            errorMessage = string.Empty;
            return true;
        }

        public bool IsReadOnly
        {
            set => _decoratedField.IsReadOnly = value;
            get => _decoratedField.IsReadOnly;
        }
        public bool EncapsulateFlag
        {
            set => _decoratedField.EncapsulateFlag = value;
            get => _decoratedField.EncapsulateFlag;
        }
        public string FieldIdentifier
        {
            set => _decoratedField.FieldIdentifier = value;
            get => _decoratedField.FieldIdentifier;
        }
        public bool CanBeReadWrite
        {
            set => _decoratedField.CanBeReadWrite = value;
            get => _decoratedField.CanBeReadWrite;
        }
        public bool HasValidEncapsulationAttributes => true;

        public QualifiedModuleName QualifiedModuleName 
            => _decoratedField.QualifiedModuleName;

        public string PropertyName
        {
            set => _decoratedField.PropertyName = value;
            get => _decoratedField.PropertyName;
        }
        public string AsTypeName_Field
        {
            set => _decoratedField.AsTypeName_Field = value;
            get => _decoratedField.AsTypeName_Field;
        }
        public string AsTypeName_Property
        {
            set => _decoratedField.AsTypeName_Property = value;
            get => _decoratedField.AsTypeName_Property;
        }
        public string ParameterName => _decoratedField.ParameterName;
        public bool ImplementLet
        {
            set => _decoratedField.ImplementLet = value;
            get => _decoratedField.ImplementLet;
        }
        public bool ImplementSet
        {
            set => _decoratedField.ImplementSet = value;
            get => _decoratedField.ImplementSet;
        }
        public IEnumerable<IPropertyGeneratorAttributes> PropertyAttributeSets => _decoratedField.PropertyAttributeSets;
        public string AsUDTMemberDeclaration { get; }

        public IEnumerable<KeyValuePair<IdentifierReference, (ParserRuleContext, string)>> ReferenceReplacements => _decoratedField.ReferenceReplacements;
        public string ReferenceWithinNewProperty => $"{Parent.ReferenceWithinNewProperty}.{_decoratedField.IdentifierName}";

        public AccessorTokens PropertyAccessor
        {
            set => _decoratedField.PropertyAccessor = value;
            get => _decoratedField.PropertyAccessor;
        }

        public AccessorTokens ReferenceAccessor
        {
            set => _decoratedField.ReferenceAccessor = value;
            get => _decoratedField.ReferenceAccessor;
        }
    }
}
