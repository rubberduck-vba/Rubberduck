using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Common;
using Rubberduck.VBEditor;
using System;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldRefactoringElement
    {
        string IdentifierName { get; }
        QualifiedModuleName QualifiedModuleName { get; }
        string AsTypeName { get; }
    }

    public interface IEncapsulateFieldCandidate : IEncapsulateFieldRefactoringElement
    {
        string TargetID { get; }
        Declaration Declaration { get; }
        bool EncapsulateFlag { get; set; }
        string BackingIdentifier { get; }
        Action<string> BackingIdentifierMutator { get; }
        string PropertyIdentifier { set; get; }
        string PropertyAsTypeName { get; }
        bool CanBeReadWrite { get; }
        bool IsReadOnly { set; get; }
        bool IsAssignedExternally { get; }
        IEncapsulateFieldConflictFinder ConflictFinder { set; get; }
        bool TryValidateEncapsulationAttributes(out string errorMessage);
    }

    public class EncapsulateFieldCandidate : IEncapsulateFieldCandidate
    {
        protected readonly int _hashCode;
        protected EncapsulationIdentifiers _fieldAndProperty;

        public EncapsulateFieldCandidate(Declaration declaration)
        {
            Declaration = declaration;
            AsTypeName = declaration.AsTypeName;

            _fieldAndProperty = new EncapsulationIdentifiers(declaration.IdentifierName);
            BackingIdentifierMutator = (value) => _fieldAndProperty.Field = value;

            IdentifierName = declaration.IdentifierName;

            TargetID = IdentifierName;

            QualifiedModuleName = declaration.QualifiedModuleName;

            //5.3.1 The declared type of a function declaration may not be a private enum.
            PropertyAsTypeName = declaration.IsEnumField() && declaration.AsTypeDeclaration.HasPrivateAccessibility()
                ? Tokens.Long
                : declaration.AsTypeName;

            CanBeReadWrite = !Declaration.IsArray;

            _hashCode = $"{QualifiedModuleName.Name}.{declaration.IdentifierName}".GetHashCode();
        }

        public Declaration Declaration { get; }

        public string IdentifierName { get; }

        public string AsTypeName { get; }

        public bool CanBeReadWrite { get; }

        private bool _isReadOnly;
        public virtual bool IsReadOnly 
        {
            set
            {
                _isReadOnly = value
                    ? !IsAssignedExternally
                    : !CanBeReadWrite;
            }
            get => _isReadOnly;
        }

        public bool IsAssignedExternally
            => Declaration.References.Any(rf => rf.IsAssignment && rf.QualifiedModuleName != Declaration.QualifiedModuleName);

        public virtual IEncapsulateFieldConflictFinder ConflictFinder { set; get; }

        public string PropertyAsTypeName { get; set; }

        public QualifiedModuleName QualifiedModuleName { get; }

        public virtual bool TryValidateEncapsulationAttributes(out string errorMessage)
        {
            (bool IsValid, string ErrorMsg) = ConflictFinder?.ValidateEncapsulationAttributes(this) ?? (true, string.Empty);

            errorMessage = ErrorMsg;
            return IsValid;
        }

        public virtual string TargetID { get; }

        protected bool _encapsulateFlag;
        public virtual bool EncapsulateFlag
        {
            set
            {
                if (_encapsulateFlag != value)
                {
                    _encapsulateFlag = value;
                    if (!_encapsulateFlag)
                    {
                        PropertyIdentifier = _fieldAndProperty.DefaultPropertyName;
                        return;
                    }

                    ConflictFinder?.AssignNoConflictIdentifiers(this);
                }
            }
            get => _encapsulateFlag;
        }

        public string PropertyIdentifier
        {
            get => _fieldAndProperty.Property;
            set
            {
                if (_fieldAndProperty.Property != value)
                {
                    _fieldAndProperty.Property = value;

                    //Reset the backing field identifier
                    _fieldAndProperty.Field = _fieldAndProperty.TargetFieldName;
                    ConflictFinder?.AssignNoConflictBackingFieldIdentifier(this);
                }
            }
        }

        public virtual string BackingIdentifier => _fieldAndProperty.Field;

        public virtual Action<string> BackingIdentifierMutator { get; }

        public override bool Equals(object obj)
        {
            return obj != null
                && obj is IEncapsulateFieldCandidate efc
                && efc.QualifiedModuleName == QualifiedModuleName
                && efc.IdentifierName == IdentifierName;
        }

        public override int GetHashCode() => _hashCode;

        public override string ToString()
            => $"({TargetID}){Declaration.ToString()}";
    }
}
