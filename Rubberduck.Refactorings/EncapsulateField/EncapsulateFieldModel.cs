using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Common;
using Rubberduck.SmartIndenter;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldModel : IRefactoringModel
    {
        private const string DEFAULT_ENCAPSULATION_UDT_IDENTIFIER = "This_Type";
        private const string DEFAULT_ENCAPSULATION_UDT_FIELD_IDENTIFIER = "this";

        private readonly IIndenter _indenter;
        private readonly IEncapsulateFieldNamesValidator _validator;

        private IDictionary<Declaration, (Declaration, IEnumerable<Declaration>)> _udtFieldToUdtDeclarationMap = new Dictionary<Declaration, (Declaration, IEnumerable<Declaration>)>();
        private IEnumerable<Declaration> UdtFields => _udtFieldToUdtDeclarationMap.Keys;
        private IEnumerable<Declaration> UdtFieldMembers(Declaration udtField) => _udtFieldToUdtDeclarationMap[udtField].Item2;

        private Dictionary<string, IEncapsulatedFieldDeclaration> _encapsulateFieldDeclarations = new Dictionary<string, IEncapsulatedFieldDeclaration>();

        public EncapsulateFieldModel(Declaration target, IEnumerable<Declaration> allMemberFields, IDictionary<Declaration, (Declaration, IEnumerable<Declaration>)> udtFieldToUdtDeclarationMap, IIndenter indenter, IEncapsulateFieldNamesValidator validator)
        {
            _indenter = indenter;
            _validator = validator;
            _udtFieldToUdtDeclarationMap = udtFieldToUdtDeclarationMap;

            foreach (var field in allMemberFields.Except(UdtFields))
            {
                var efd = EncapsulateDeclaration(field);
                _encapsulateFieldDeclarations.Add(efd.TargetID, efd);
            }

            AddUDTEncapsulationFields(udtFieldToUdtDeclarationMap);

            this[target].EncapsulationAttributes.EncapsulateFlag = true;
        }

        public IEnumerable<IEncapsulatedFieldDeclaration> FlaggedEncapsulationFields => _encapsulateFieldDeclarations.Values.Where(v => v.EncapsulationAttributes.EncapsulateFlag);

        public IEnumerable<string> EncapsulationFieldIDs
            => _encapsulateFieldDeclarations.Keys;

        public IEnumerable<IEncapsulatedFieldDeclaration> EncapsulationFields
            => _encapsulateFieldDeclarations.Values;

        public IEncapsulatedFieldDeclaration this[string encapsulatedFieldIdentifier] 
            => _encapsulateFieldDeclarations[encapsulatedFieldIdentifier];

        public IEncapsulatedFieldDeclaration this[Declaration fieldDeclaration] 
            => _encapsulateFieldDeclarations.Values.Where(efd => efd.Declaration.Equals(fieldDeclaration)).Select(encapsulatedField => encapsulatedField).Single();

        public bool EncapsulateWithUDT { set; get; }

        public string EncapsulateWithUDT_TypeIdentifier { set; get; } = DEFAULT_ENCAPSULATION_UDT_IDENTIFIER;

        public string EncapsulateWithUDT_FieldName { set; get; } = DEFAULT_ENCAPSULATION_UDT_FIELD_IDENTIFIER;

        public IList<string> PropertiesContent
        {
            get
            {
                var textBlocks = new List<string>();
                foreach (var field in FlaggedEncapsulationFields)
                {
                    textBlocks.Add(BuildPropertiesTextBlock(field.EncapsulationAttributes));
                }
                return textBlocks;
            }
        }

        public IEncapsulateFieldNewContentProvider NewContent
        {
            get
            {
                var newContent = new EncapsulateFieldNewContent();
                newContent = LoadNewDeclarationsContent(newContent);
                newContent = LoadNewPropertiesContent(newContent);
                return newContent;
            }
        }

        public EncapsulateFieldNewContent LoadNewPropertiesContent(EncapsulateFieldNewContent newContent)
        {
            if (!FlaggedEncapsulationFields.Any()) { return newContent; }

            newContent.AddCodeBlock($"{string.Join($"{Environment.NewLine}{Environment.NewLine}", PropertiesContent)}");
            return newContent;
        }

        public EncapsulateFieldNewContent LoadNewDeclarationsContent(EncapsulateFieldNewContent newContent)
        {
            var nonUdtMemberFields = FlaggedEncapsulationFields
                    .Where(encFld => encFld.Declaration.IsVariable());

            if (EncapsulateWithUDT)
            {
                var udt = new UDTGenerator(EncapsulateWithUDT_TypeIdentifier, _indenter);
                foreach (var nonUdtMemberField in nonUdtMemberFields)
                {
                    udt.AddMember(nonUdtMemberField);
                }
                newContent.AddDeclarationBlock(udt.TypeDeclarationBlock);
                newContent.AddDeclarationBlock(udt.FieldDeclaration(EncapsulateWithUDT_FieldName));

                //TODO: handle selected UDTs
                return newContent;
            }

            foreach (var nonUdtMemberField in nonUdtMemberFields)
            {
                var attributes = nonUdtMemberField.EncapsulationAttributes;

                if (nonUdtMemberField.Declaration.Accessibility == Accessibility.Private && attributes.NewFieldName.Equals(nonUdtMemberField.Declaration.IdentifierName))
                {
                    continue;
                }

                if (nonUdtMemberField.Declaration.IsDeclaredInList())
                {
                    var targetIdentifier = nonUdtMemberField.Declaration.Context.GetText().Replace(attributes.TargetName, attributes.NewFieldName);
                    var newField = nonUdtMemberField.Declaration.IsTypeSpecified ? $"{Tokens.Private} {targetIdentifier}" : $"{Tokens.Private} {targetIdentifier} {Tokens.As} {nonUdtMemberField.Declaration.AsTypeName}";

                    newContent.AddDeclarationBlock(newField);
                }
            }
            return newContent;
        }

        private void AddUDTEncapsulationFields(IDictionary<Declaration, (Declaration, IEnumerable<Declaration>)> udtFieldToTypeMap)
        {
            foreach (var udtField in udtFieldToTypeMap.Keys)
            {
                var udtEncapsulation = DecorateUDTVariableDeclaration(udtField);
                _encapsulateFieldDeclarations.Add(udtEncapsulation.TargetID, udtEncapsulation);


                foreach (var udtMember in UdtFieldMembers(udtField))
                {
                    var encapsulatedUdtMember = EncapsulateDeclaration(udtMember);
                    encapsulatedUdtMember = DecorateUDTMember(encapsulatedUdtMember, udtEncapsulation as EncapsulatedUserDefinedType);
                    _encapsulateFieldDeclarations.Add(encapsulatedUdtMember.TargetID, encapsulatedUdtMember);
                }
            }
        }

        private IEncapsulatedFieldDeclaration EncapsulateDeclaration(Declaration target)
        {
            var encapsulated = new EncapsulatedFieldDeclaration(target, _validator);
            if (target.IsArray)
            {
                return EncapsulatedArrayType.Decorate(encapsulated);
            }
            else if (target.AsTypeName.Equals(Tokens.Variant))
            {
                return EncapsulatedVariantType.Decorate(encapsulated);
            }
            else if (target.IsObject)
            {
                return EncapsulatedObjectType.Decorate(encapsulated);
            }
            return EncapsulatedValueType.Decorate(encapsulated);
        }

        private IEncapsulatedFieldDeclaration DecorateUDTVariableDeclaration(Declaration target)
        {
            return EncapsulatedUserDefinedType.Decorate(new EncapsulatedFieldDeclaration(target, _validator));
        }

        private IEncapsulatedFieldDeclaration DecorateUDTMember(IEncapsulatedFieldDeclaration udtMember, EncapsulatedUserDefinedType udtVariable)
        {
            var targetIDPair = new KeyValuePair<Declaration, string>(udtMember.Declaration,$"{udtVariable.Declaration.IdentifierName}.{udtMember.Declaration.IdentifierName}");
            return EncapsulatedUserDefinedTypeMember.Decorate(udtMember, udtVariable, HasMultipleInstantiationsOfSameType(udtVariable.Declaration, targetIDPair));
        }

        private bool HasMultipleInstantiationsOfSameType(Declaration udtVariable, KeyValuePair<Declaration, string> targetIDPair)
        {
            var udt = _udtFieldToUdtDeclarationMap[udtVariable].Item1;
            var otherVariableOfTheSameType = _udtFieldToUdtDeclarationMap.Keys.Where(k => k != udtVariable && _udtFieldToUdtDeclarationMap[k].Item1 == udt);
            return otherVariableOfTheSameType.Any();
        }

        private string BuildPropertiesTextBlock(IFieldEncapsulationAttributes attributes)
        {
            var generator = new PropertyGenerator
            {
                PropertyName = attributes.PropertyName,
                AsTypeName = attributes.AsTypeName,
                BackingField = EncapsulateWithUDT 
                                    ? $"{EncapsulateWithUDT_FieldName}.{attributes.PropertyName}" 
                                    : attributes.FieldReferenceExpression,
                ParameterName = attributes.ParameterName,
                GenerateSetter = attributes.ImplementSetSetterType,
                GenerateLetter = attributes.ImplementLetSetterType
            };

            var propertyTextLines = generator.AllPropertyCode.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            return string.Join(Environment.NewLine, _indenter.Indent(propertyTextLines, true));
        }
    }
}
