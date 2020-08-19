using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Common;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Resources;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IObjectStateUDT : IEncapsulateFieldRefactoringElement
    {
        string TypeIdentifier { set; get; }
        string FieldIdentifier { set; get; }
        string TypeDeclarationBlock(IIndenter indenter = null);
        string FieldDeclarationBlock { get; }
        void AddMembers(IEnumerable<IConvertToUDTMember> fields);
        bool IsExistingDeclaration { get; }
        Declaration AsTypeDeclaration { get; }
        bool IsSelected { set; get; }
        IEnumerable<IUserDefinedTypeMemberCandidate> ExistingMembers { get; }
    }

    //ObjectStateUDT can be an existing UDT (Private only) selected by the user, or a 
    //newly inserted declaration
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
        }

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
                }

                if (_isSelected && IsExistingDeclaration)
                {
                    _wrappedUDT.EncapsulateFlag = false;
                }
            }
            get => _isSelected;
        }

        public IEnumerable<IUserDefinedTypeMemberCandidate> ExistingMembers
        {
            get
            {
                if (IsExistingDeclaration)
                {
                    return _wrappedUDT.Members;
                }
                return Enumerable.Empty<IUserDefinedTypeMemberCandidate>();
            }
        }

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

        public void AddMembers(IEnumerable<IConvertToUDTMember> fields)
        {
            _convertedMembers = new List<IConvertToUDTMember>();
            if (IsExistingDeclaration)
            {
                foreach (var member in _wrappedUDT.Members)
                {
                    var convertedMember = new ConvertToUDTMember(member, this) { EncapsulateFlag = false };
                    _convertedMembers.Add(convertedMember);
                }
            }
            _convertedMembers.AddRange(fields);
        }

        public string FieldDeclarationBlock
            => $"{Accessibility.Private} {IdentifierName} {Tokens.As} {AsTypeName}";

        public string TypeDeclarationBlock(IIndenter indenter = null)
        {
            var udtMembers = _convertedMembers.Where(m => m.Declaration is VariableDeclaration)
                .Select(m => (m.Declaration as VariableDeclaration, m.BackingIdentifier));

            return _codeBuilder.BuildUserDefinedTypeDeclaration(AsTypeName, udtMembers);
        }

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
