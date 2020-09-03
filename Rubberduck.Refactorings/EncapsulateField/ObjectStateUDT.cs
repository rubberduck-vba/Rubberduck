using Rubberduck.Parsing.Symbols;
using Rubberduck.Common;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Resources;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IObjectStateUDT : IEncapsulateFieldRefactoringElement
    {
        string TypeIdentifier { set; get; }
        string FieldDeclarationBlock { get; }
        string FieldIdentifier { set; get; }
        bool IsExistingDeclaration { get; }
        Declaration AsTypeDeclaration { get; }
        bool IsSelected { set; get; }
        IEnumerable<IUserDefinedTypeMemberCandidate> ExistingMembers { get; }
    }

    /// <summary>
    /// ObjectStateUDT is a Private UserDefinedType whose UserDefinedTypeMembers represent
    /// object state in lieu of (or in addition to) a set of Private fields.
    /// </summary>
    /// <remarks>
    /// Within the EncapsulateField refactoring, the ObjectStateUDT can be an existing
    /// UserDefinedType or an identifier that will be used to generate a new UserDefinedType
    /// </remarks>
    public class ObjectStateUDT : IObjectStateUDT
    {
        private static string _defaultNewFieldName = RubberduckUI.EncapsulateField_DefaultObjectStateUDTFieldName;
        private List<IConvertToUDTMember> _convertedMembers;

        private readonly IUserDefinedTypeCandidate _wrappedUDT;
        private readonly ICodeBuilder _codeBuilder;
        private int _hashCode;

        public ObjectStateUDT(IUserDefinedTypeCandidate udt)
            : this(udt.Declaration.AsTypeName)
        {
            if (!udt.TypeDeclarationIsPrivate)
            {
                throw new ArgumentException();
            }

            FieldIdentifier = udt.IdentifierName;
            _wrappedUDT = udt;
            _hashCode = ($"{_qmn.Name}.{_wrappedUDT.IdentifierName}").GetHashCode();
        }

        public ObjectStateUDT(QualifiedModuleName qmn)
            :this($"T{qmn.ComponentName.CapitalizeFirstLetter()}")
        {
            QualifiedModuleName = qmn;
        }

        private ObjectStateUDT(string typeIdentifier)
        {
            FieldIdentifier = _defaultNewFieldName;
            TypeIdentifier = typeIdentifier;
            _convertedMembers = new List<IConvertToUDTMember>();
            _codeBuilder = new CodeBuilder();
            _convertedMembers = new List<IConvertToUDTMember>();
        }

        public string FieldDeclarationBlock
            => $"{Accessibility.Private} {IdentifierName} {Tokens.As} {AsTypeName}";

        public string IdentifierName => _wrappedUDT?.IdentifierName ?? FieldIdentifier;

        public string AsTypeName => _wrappedUDT?.AsTypeName ?? TypeIdentifier;

        private bool _isSelected;
        public bool IsSelected
        {
            set
            {
                _isSelected = value;
                if (_wrappedUDT != null)
                {
                    _wrappedUDT.IsSelectedObjectStateUDT = value;
                    if (_isSelected && IsExistingDeclaration)
                    {
                        _wrappedUDT.EncapsulateFlag = false;
                    }
                }
            }
            get => _isSelected;
        }

        public IEnumerable<IUserDefinedTypeMemberCandidate> ExistingMembers 
            => IsExistingDeclaration
                ? _wrappedUDT.Members
                : Enumerable.Empty<IUserDefinedTypeMemberCandidate>();

        private QualifiedModuleName _qmn;
        public QualifiedModuleName QualifiedModuleName
        {
            set => _qmn = value;
            get => _wrappedUDT?.QualifiedModuleName ?? _qmn;
        }

        public string TypeIdentifier { set; get; }

        public bool IsExistingDeclaration => _wrappedUDT != null;

        public Declaration AsTypeDeclaration => _wrappedUDT?.Declaration.AsTypeDeclaration;

        public string FieldIdentifier { set; get; }

        public override bool Equals(object obj)
        {
            if (obj is IObjectStateUDT stateUDT && stateUDT.FieldIdentifier == FieldIdentifier)
            {
                return true;
            }
            if (obj is IEncapsulateFieldRefactoringElement fd && fd.IdentifierName == IdentifierName)
            {
                return true;
            }
            return false;
        }

       public override int GetHashCode() => _hashCode;
    }
}
