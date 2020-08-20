using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IUserDefinedTypeMemberCandidate : IEncapsulateFieldCandidate
    {
        IUserDefinedTypeCandidate UDTField { get; }
        PropertyAttributeSet AsPropertyGeneratorSpec { get; }
        IEnumerable<IdentifierReference> FieldContextReferences { get; }
        IEncapsulateFieldCandidate WrappedCandidate { get; }
    }

    public class UserDefinedTypeMemberCandidate : IUserDefinedTypeMemberCandidate
    {
        private int _hashCode;
        private readonly string _uniqueID;
        private string _rhsParameterIdentifierName;
        public UserDefinedTypeMemberCandidate(IEncapsulateFieldCandidate candidate, IUserDefinedTypeCandidate udtField)
        {
            _wrappedCandidate = candidate;
            _rhsParameterIdentifierName = Resources.Refactorings.Refactorings.CodeBuilder_DefaultPropertyRHSParam;
            UDTField = udtField;
            PropertyIdentifier = IdentifierName;
            BackingIdentifier = IdentifierName;
            _uniqueID = BuildUniqueID(candidate, UDTField);
            _hashCode = _uniqueID.GetHashCode();
        }

        private IEncapsulateFieldCandidate _wrappedCandidate;

        public IEncapsulateFieldCandidate WrappedCandidate => _wrappedCandidate;

        public string AsTypeName => _wrappedCandidate.AsTypeName;

        public string BackingIdentifier
        {
            get
            {
                  return _wrappedCandidate.IdentifierName;
            }
            set { }
        }

        public string BackingAsTypeName => Declaration.AsTypeName;

        public IUserDefinedTypeCandidate UDTField { private set; get; }

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

        public string TargetID => $"{UDTField.IdentifierName}.{IdentifierName}";

        public IEnumerable<IdentifierReference> FieldContextReferences
            => GetUDTMemberReferencesForField(this, UDTField);

        public string IdentifierForReference(IdentifierReference idRef)
            => PropertyIdentifier;

        public PropertyAttributeSet AsPropertyGeneratorSpec
        {
            get
            {
                return new PropertyAttributeSet()
                {
                    PropertyName = PropertyIdentifier,
                    BackingField = BackingIdentifier,
                    AsTypeName = PropertyAsTypeName,
                    ParameterName = ParameterName,
                    GenerateLetter = ImplementLet,
                    GenerateSetter = ImplementSet,
                    UsesSetAssignment = Declaration.IsObject,
                    IsUDTProperty = Declaration.DeclarationType == DeclarationType.UserDefinedType,
                    Declaration = Declaration
                };
            }
        }

        public override bool Equals(object obj)
        {
            return obj != null
                && obj is IUserDefinedTypeMemberCandidate udtMember
                && BuildUniqueID(udtMember, udtMember.UDTField) == _uniqueID;
        }

        public override int GetHashCode() => _hashCode;

        public string PropertyIdentifier { set; get; }

        private static string BuildUniqueID(IEncapsulateFieldCandidate candidate, IEncapsulateFieldCandidate field) => $"{candidate.QualifiedModuleName.Name}.{field.IdentifierName}.{candidate.IdentifierName}";

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
                if (_wrappedCandidate is IUserDefinedTypeCandidate udt && udt.TypeDeclarationIsPrivate)
                {
                    foreach (var member in udt.Members)
                    {
                        member.EncapsulateFlag = value;
                    }
                    return;
                }
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

        public string ParameterName => _rhsParameterIdentifierName;

        public bool ImplementLet => _wrappedCandidate.ImplementLet;

        public bool ImplementSet => _wrappedCandidate.ImplementSet;

        public IEnumerable<PropertyAttributeSet> PropertyAttributeSets
        {
            get
            {
                if (!(_wrappedCandidate is IUserDefinedTypeCandidate udt))
                {
                    return new List<PropertyAttributeSet>() { AsPropertyGeneratorSpec };
                }

                var sets = _wrappedCandidate.PropertyAttributeSets;
                if (udt.TypeDeclarationIsPrivate)
                {
                    return sets;
                }
                var modifiedSets = new List<PropertyAttributeSet>();
                for(var idx = 0; idx < sets.Count(); idx++)
                {
                    var attr = sets.ElementAt(idx);
                    attr.BackingField = attr.PropertyName;
                    modifiedSets.Add(attr);
                }
                return modifiedSets;
            }
        }
    }
}
