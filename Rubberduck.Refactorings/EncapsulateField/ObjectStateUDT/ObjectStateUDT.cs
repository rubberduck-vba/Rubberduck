﻿using Rubberduck.Parsing.Symbols;
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
        Declaration Declaration { get; }
        string TypeIdentifier { set; get; }
        string FieldDeclarationBlock { get; }
        string FieldIdentifier { set; get; }
        bool IsExistingDeclaration { get; }
        Declaration AsTypeDeclaration { get; }
        bool IsSelected { set; get; }
        IReadOnlyCollection<IUserDefinedTypeMemberCandidate> ExistingMembers { get; }
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
        private readonly IUserDefinedTypeCandidate _wrappedUDTField;
        private readonly int _hashCode;

        public ObjectStateUDT(IUserDefinedTypeCandidate udtField)
            : this(udtField.Declaration.AsTypeName)
        {
            if (!udtField.TypeDeclarationIsPrivate)
            {
                throw new ArgumentException();
            }

            QualifiedModuleName = udtField.QualifiedModuleName;
            FieldIdentifier = udtField.IdentifierName;
            _wrappedUDTField = udtField;
            _hashCode = ($"{QualifiedModuleName.Name}.{_wrappedUDTField.IdentifierName}").GetHashCode();
        }

        public ObjectStateUDT(QualifiedModuleName qualifiedModuleName)
            :this($"T{qualifiedModuleName.ComponentName.CapitalizeFirstLetter()}")
        {
            QualifiedModuleName = qualifiedModuleName;
        }

        private ObjectStateUDT(string typeIdentifier)
        {
            FieldIdentifier = _defaultNewFieldName;
            TypeIdentifier = typeIdentifier;
        }

        public string FieldDeclarationBlock
            => $"{Accessibility.Private} {IdentifierName} {Tokens.As} {AsTypeName}";

        public string IdentifierName => _wrappedUDTField?.IdentifierName ?? FieldIdentifier;

        public Declaration Declaration => _wrappedUDTField?.Declaration;

        public string AsTypeName => _wrappedUDTField?.AsTypeName ?? TypeIdentifier;

        private bool _isSelected;
        public bool IsSelected
        {
            set
            {
                _isSelected = value;
                if (_wrappedUDTField != null)
                {
                    _wrappedUDTField.IsSelectedObjectStateUDT = value;
                    if (_isSelected && IsExistingDeclaration)
                    {
                        _wrappedUDTField.EncapsulateFlag = false;
                    }
                }
            }
            get => _isSelected;
        }

        public IReadOnlyCollection<IUserDefinedTypeMemberCandidate> ExistingMembers 
            => _wrappedUDTField?.Members.ToList() ?? new List<IUserDefinedTypeMemberCandidate>();

        public QualifiedModuleName QualifiedModuleName { get; }

        public string TypeIdentifier { set; get; }

        public bool IsExistingDeclaration => _wrappedUDTField != null;

        public Declaration AsTypeDeclaration => _wrappedUDTField?.Declaration.AsTypeDeclaration;

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