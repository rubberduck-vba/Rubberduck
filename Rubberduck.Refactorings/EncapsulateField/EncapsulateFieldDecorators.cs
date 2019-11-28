using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Common;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Windows;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulatedUserDefinedTypeMember : IEncapsulateFieldCandidate //: EncapsulateFieldDecoratorBase
    {
        private IEncapsulateFieldCandidate _candidate;
        private IFieldEncapsulationAttributes _udtVariableAttributes;
        private bool _nameResolveProperty;
        private string _originalVariableName;
        private string _targetID;
        public EncapsulatedUserDefinedTypeMember(IEncapsulateFieldCandidate candidate, IEncapsulateFieldCandidate udtVariable, bool propertyIdentifierRequiresNameResolution)
        {
            _candidate = candidate;
            _originalVariableName = udtVariable.Declaration.IdentifierName;
            _nameResolveProperty = propertyIdentifierRequiresNameResolution;
            _udtVariableAttributes = udtVariable.EncapsulationAttributes;

            EncapsulationAttributes.PropertyName = BuildPropertyName();
            if (EncapsulationAttributes is FieldEncapsulationAttributes fea)
            {
                fea.FieldReferenceExpressionFunc =
                 () =>  { var prefix = _udtVariableAttributes.EncapsulateFlag
                                         ? _udtVariableAttributes.NewFieldName
                                         : _udtVariableAttributes.TargetName;

                            return $"{prefix}.{EncapsulationAttributes.NewFieldName}";
                        };
            }

            _targetID = $"{udtVariable.Declaration.IdentifierName}.{Declaration.IdentifierName}";
            candidate.IsUDTMember = true;
        }

        public Declaration Declaration => _candidate.Declaration;

        public DeclarationType DeclarationType => _candidate.DeclarationType;

        public IFieldEncapsulationAttributes EncapsulationAttributes
        {
            get => _candidate.EncapsulationAttributes;
            set => _candidate.EncapsulationAttributes = value;
        }

        public bool IsReadOnly
        {
            get => _candidate.EncapsulationAttributes.ReadOnly;
            set => _candidate.EncapsulationAttributes.ReadOnly = value;
        }

        public bool EncapsulateFlag
        {
            get => _candidate.EncapsulateFlag;
            set => _candidate.EncapsulateFlag = value;
        }

        public bool CanBeReadWrite
        {
            get => _candidate.CanBeReadWrite;
            set => _candidate.CanBeReadWrite = value;
        }

        public string PropertyName
        {
            get => _candidate.EncapsulationAttributes.PropertyName;
            set => _candidate.EncapsulationAttributes.PropertyName = value;
        }

        public string NewFieldName
        {
            get => _candidate.EncapsulationAttributes.NewFieldName;
        }

        public string AsTypeName => _candidate.EncapsulationAttributes.AsTypeName;

        public bool IsUDTMember
        {
            get => _candidate.IsUDTMember;
            set => _candidate.IsUDTMember = value;
        }

        public bool HasValidEncapsulationAttributes
            => _candidate.HasValidEncapsulationAttributes;

        public QualifiedModuleName QualifiedModuleName => _candidate.QualifiedModuleName;

        public IEnumerable<IdentifierReference> References => _candidate.References;

        public string TargetID => _targetID;

        private string BuildPropertyName()
        {
            if (_nameResolveProperty)
            {
                var propertyPrefix = char.ToUpper(_originalVariableName[0]) + _originalVariableName.Substring(1, _originalVariableName.Length - 1);
                return $"{propertyPrefix}_{EncapsulationAttributes.TargetName}";
            }
            return EncapsulationAttributes.TargetName;
        }
    }
}
