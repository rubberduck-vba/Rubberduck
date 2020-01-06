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
        IEnumerable<IdentifierReference> ParentContextReferences { get; }
        void LoadReferenceExpressions();
        bool IsExistingMember { get; }
    }

    public class UserDefinedTypeMemberCandidate : IUserDefinedTypeMemberCandidate, IConvertToUDTMember
    {
        private int _hashCode;
        private readonly string _uniqueID;
        public UserDefinedTypeMemberCandidate(IEncapsulateFieldCandidate candidate, IUserDefinedTypeCandidate udtVariable/*, IValidateEncapsulateFieldNames validator*/)
        {
            _wrappedCandidate = candidate;
            Parent = udtVariable;
            PropertyIdentifier = IdentifierName;
            _uniqueID = BuildUniqueID(candidate);
            _hashCode = _uniqueID.GetHashCode();
        }

        private IEncapsulateFieldCandidate _wrappedCandidate;

        public string AsTypeName => _wrappedCandidate.AsTypeName;

        //public string AccessorInProperty { /*set;*/ get; }

        public IUserDefinedTypeCandidate Parent { private set; get; }

        public void LoadFieldReferenceContextReplacements(string referenceQualifier = null)
        {
            ReferenceQualifier = referenceQualifier;
        }

        public IValidateVBAIdentifiers NameValidator
        {
            set => _wrappedCandidate.NameValidator = value;
            get => _wrappedCandidate.NameValidator;
        }

        public IEncapsulateFieldConflictFinder ConflictFinder
        {
            set => _wrappedCandidate.ConflictFinder = value;
            get => _wrappedCandidate.ConflictFinder;
        }

        public string ReferenceQualifier { set; get; }

        public string TargetID => $"{Parent.IdentifierName}.{IdentifierName}";

        public IEnumerable<IdentifierReference> ParentContextReferences
            => GetUDTMemberReferencesForField(this, Parent);

        public void LoadReferenceExpressions()
        {
            foreach (var rf in ParentContextReferences)
            {
                if (rf.QualifiedModuleName == QualifiedModuleName
                    && !rf.Context.TryGetAncestor<VBAParser.WithMemberAccessExprContext>(out _))
                {
                    SetReferenceRewriteContent(rf, PropertyIdentifier);
                    continue;
                }
                var moduleQualifier = rf.Context.TryGetAncestor<VBAParser.WithStmtContext>(out _)
                    || rf.QualifiedModuleName == QualifiedModuleName
                    ? string.Empty
                    : $"{QualifiedModuleName.ComponentName}";

                SetReferenceRewriteContent(rf, $"{moduleQualifier}.{PropertyIdentifier}");
            }
        }

        protected void SetReferenceRewriteContent(IdentifierReference idRef, string replacementText)
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
                    PropertyName = PropertyIdentifier,
                    BackingField = ReferenceWithinNewProperty,
                    AsTypeName = PropertyAsTypeName,
                    ParameterName = ParameterName,
                    GenerateLetter = ImplementLet,
                    GenerateSetter = ImplementSet,
                    UsesSetAssignment = Declaration.IsObject,
                    IsUDTProperty = false //TODO: If the member is a UDT, this needs to be true
                };
            }
        }

        public Dictionary<IdentifierReference, (ParserRuleContext, string)> IdentifierReplacements { get; } = new Dictionary<IdentifierReference, (ParserRuleContext, string)>();

        public override bool Equals(object obj)
        {
            return obj != null
                && obj is IUserDefinedTypeMemberCandidate udtMember
                && BuildUniqueID(udtMember) == _uniqueID;
        }

        public override int GetHashCode() => _hashCode;

        /*
         * IConvertToUDTMemberInterface
         */
        public string AccessorInProperty
            => $"{Parent.FieldIdentifier}.{Declaration.IdentifierName}";

        public string AccessorLocalReference
            => $"{Parent.FieldIdentifier}.{PropertyIdentifier}";

        public string AccessorExternalReference { set; get; }
        public string PropertyIdentifier { set; get; }

        public string PropertyAsTypeName2 { set; get; }
        public string UDTMemberIdentifier { set; get; }

        public virtual string UDTMemberDeclaration
            => $"{Declaration.IdentifierName} {Tokens.As} {FieldAsTypeName}";

        public IObjectStateUDT ObjectStateUDT { set; get; }

        private static string BuildUniqueID(IEncapsulateFieldCandidate candidate) => $"{candidate.QualifiedModuleName.Name}.{candidate.IdentifierName}";

        private static IEnumerable<IdentifierReference> GetUDTMemberReferencesForField(IEncapsulateFieldCandidate udtMember, IUserDefinedTypeCandidate field)
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

        public Declaration Declaration => _wrappedCandidate.Declaration;
        public string IdentifierName => _wrappedCandidate.IdentifierName;

        public bool TryValidateEncapsulationAttributes(out string errorMessage)
        {
            errorMessage = string.Empty;
            return true;
        }

        public bool IsReadOnly
        {
            set => _wrappedCandidate.IsReadOnly = value;
            get => _wrappedCandidate.IsReadOnly;
        }

        private bool _encapsulateFlag;
        public bool EncapsulateFlag
        {
            set
            {
                _encapsulateFlag = value;
                if (!IsExistingMember)
                {
                    _wrappedCandidate.EncapsulateFlag = value;
                }
            }

            get => _encapsulateFlag;
        }

        public bool IsExistingMember => _wrappedCandidate.Declaration.ParentDeclaration.DeclarationType is DeclarationType.UserDefinedType;

        public string FieldIdentifier
        {
            set => _wrappedCandidate.FieldIdentifier = value;
            get => _wrappedCandidate.FieldIdentifier;
        }

        public bool CanBeReadWrite
        {
            set => _wrappedCandidate.CanBeReadWrite = value;
            get => _wrappedCandidate.CanBeReadWrite;
        }
        public bool HasValidEncapsulationAttributes => true;

        public QualifiedModuleName QualifiedModuleName 
            => _wrappedCandidate.QualifiedModuleName;

        public string FieldAsTypeName // AsTypeName_Field
        {
            set => _wrappedCandidate.FieldAsTypeName = value;
            get => _wrappedCandidate.FieldAsTypeName;
        }
        public string PropertyAsTypeName
        {
            set => _wrappedCandidate.PropertyAsTypeName = value;
            get => _wrappedCandidate.PropertyAsTypeName;
        }
        public string ParameterName => _wrappedCandidate.ParameterName;

        public bool ImplementLet => _wrappedCandidate.ImplementLet;

        public bool ImplementSet => _wrappedCandidate.ImplementSet;

        private bool _convertFieldToUDTMember;
        public bool ConvertFieldToUDTMember
        {
            set  => _convertFieldToUDTMember = value;
            get => false;
        }

        public EncapsulateFieldStrategy EncapsulateFieldStrategy { set; get; } = EncapsulateFieldStrategy.UseBackingFields;

        public IEnumerable<IPropertyGeneratorAttributes> PropertyAttributeSets => _wrappedCandidate.PropertyAttributeSets;

        public IEnumerable<KeyValuePair<IdentifierReference, (ParserRuleContext, string)>> ReferenceReplacements => _wrappedCandidate.ReferenceReplacements;

        private string ReferenceWithinNewProperty => $"{ReferenceQualifier}.{_wrappedCandidate.IdentifierName}";
    }
}
