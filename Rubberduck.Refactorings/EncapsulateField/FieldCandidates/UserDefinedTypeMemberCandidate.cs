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
        PropertyAttributeSet AsPropertyGeneratorSpec { get; }
        IEnumerable<IdentifierReference> ParentContextReferences { get; }
    }

    public class UserDefinedTypeMemberCandidate : IUserDefinedTypeMemberCandidate
    {
        private int _hashCode;
        private readonly string _uniqueID;
        public UserDefinedTypeMemberCandidate(IEncapsulateFieldCandidate candidate, IUserDefinedTypeCandidate udtVariable)
        {
            _wrappedCandidate = candidate;
            Parent = udtVariable;
            PropertyIdentifier = IdentifierName;
            BackingIdentifier = IdentifierName;
            _uniqueID = BuildUniqueID(candidate);
            _hashCode = _uniqueID.GetHashCode();
        }

        private IEncapsulateFieldCandidate _wrappedCandidate;

        public string AsTypeName => _wrappedCandidate.AsTypeName;

        public string BackingIdentifier
        {
            get => _wrappedCandidate.IdentifierName;
            set { }
        }

        public string BackingAsTypeName => Declaration.AsTypeName;

        public IUserDefinedTypeCandidate Parent { private set; get; }

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

        public string TargetID => $"{Parent.IdentifierName}.{IdentifierName}";

        public IEnumerable<IdentifierReference> ParentContextReferences
            => GetUDTMemberReferencesForField(this, Parent);

        public string ReferenceAccessor(IdentifierReference idRef)
            => PropertyIdentifier;

        public PropertyAttributeSet AsPropertyGeneratorSpec
        {
            get
            {
                return new PropertyAttributeSet()
                {
                    PropertyName = PropertyIdentifier,
                    BackingField = $"{Parent.BackingIdentifier}.{BackingIdentifier}",
                    AsTypeName = PropertyAsTypeName,
                    ParameterName = ParameterName,
                    GenerateLetter = ImplementLet,
                    GenerateSetter = ImplementSet,
                    UsesSetAssignment = Declaration.IsObject,
                    IsUDTProperty = false //TODO: If udtMember is a UDT, this needs to be true
                };
            }
        }

        public override bool Equals(object obj)
        {
            return obj != null
                && obj is IUserDefinedTypeMemberCandidate udtMember
                && BuildUniqueID(udtMember) == _uniqueID;
        }

        public override int GetHashCode() => _hashCode;

        public string PropertyIdentifier { set; get; }

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
                var valueChanged = _encapsulateFlag != value;

                _encapsulateFlag = value;
                if (!_encapsulateFlag)
                {
                    _wrappedCandidate.EncapsulateFlag = value;
                    PropertyIdentifier = _wrappedCandidate.PropertyIdentifier;
                }
                else if (valueChanged)
                {
                    ConflictFinder.AssignNoConflictIdentifiers(this);
                }
            }
            get => _encapsulateFlag;
        }

        public bool CanBeReadWrite
        {
            set => _wrappedCandidate.CanBeReadWrite = value;
            get => _wrappedCandidate.CanBeReadWrite;
        }
        public bool HasValidEncapsulationAttributes => true;

        public QualifiedModuleName QualifiedModuleName 
            => _wrappedCandidate.QualifiedModuleName;

        public string PropertyAsTypeName => _wrappedCandidate.PropertyAsTypeName;

        public string ParameterName
        {
            set => _wrappedCandidate.ParameterName = value;
            get => _wrappedCandidate.ParameterName;
        }

    public bool ImplementLet => _wrappedCandidate.ImplementLet;

        public bool ImplementSet => _wrappedCandidate.ImplementSet;

        public IEnumerable<PropertyAttributeSet> PropertyAttributeSets => _wrappedCandidate.PropertyAttributeSets;
    }
}
