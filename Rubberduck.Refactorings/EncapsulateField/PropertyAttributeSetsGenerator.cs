using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public struct PropertyAttributeSet
    {
        public string PropertyName { get; set; }
        public string BackingField { get; set; }
        public string AsTypeName { get; set; }
        public string RHSParameterIdentifier { get; set; }
        public bool GeneratePropertyLet { get; set; }
        public bool GeneratePropertySet { get; set; }
        public bool UsesSetAssignment { get; set; }
        public bool IsUDTProperty { get; set; }
        public Declaration Declaration { get; set; }
    }

    public interface IPropertyAttributeSetsGenerator
    {
        IReadOnlyCollection<PropertyAttributeSet> GeneratePropertyAttributeSets(IEncapsulateFieldCandidate candidate);
    }

    /// <summary>
    /// PropertyAttributeSetsGenerator operates on an <c>IEncapsulateFieldCandidate</c> instance to
    /// generate a collection of <c>PropertyAttributeSet</c>s used by the EncapsulateField refactoring
    /// actions to generate Property Let/Set/Get code blocks.
    /// </summary>
    /// <remarks>
    /// Typically there is only a single <c>PropertyAttributeSet</c> in the collection.
    /// In the case of a Private UserDefinedType, there will be a <c>PropertyAttributeSet</c>
    /// for each UserDefinedTypeMember.
    /// </remarks>
    public class PropertyAttributeSetsGenerator : IPropertyAttributeSetsGenerator
    {
        private Func<IEncapsulateFieldCandidate, string, string> _backingFieldQualifierFunc;

        public PropertyAttributeSetsGenerator()
        {
            _backingFieldQualifierFunc = BackingField_BackingFieldQualifier;
        }

        private static string BackingUDTMember_BackingFieldQualifier(IEncapsulateFieldCandidate candidate, string backingField)
            => $"{candidate.PropertyIdentifier}.{backingField}";

        private static string BackingField_BackingFieldQualifier(IEncapsulateFieldCandidate candidate, string backingField)
            => $"{candidate.BackingIdentifier}.{backingField}";

        public IReadOnlyCollection<PropertyAttributeSet> GeneratePropertyAttributeSets(IEncapsulateFieldCandidate candidate)
        {
            if (!(candidate is IEncapsulateFieldAsUDTMemberCandidate asUDTCandidate))
            {
                _backingFieldQualifierFunc = BackingField_BackingFieldQualifier;
                return CreatePropertyAttributeSets(candidate).ToList();
            }

            return GeneratePropertyAttributeSets(asUDTCandidate);
        }

        private IReadOnlyCollection<PropertyAttributeSet> GeneratePropertyAttributeSets(IEncapsulateFieldAsUDTMemberCandidate asUDTCandidate)
        {
            _backingFieldQualifierFunc = BackingUDTMember_BackingFieldQualifier;

            Func<PropertyAttributeSet, string> QualifyPrivateUDTWrappedBackingField = attributeSet =>
            {
                var fields = attributeSet.BackingField.Split(new char[] { '.' });

                return fields.Count() > 1
                    ? $"{asUDTCandidate.ObjectStateUDT.FieldIdentifier}.{attributeSet.BackingField}"
                    : $"{asUDTCandidate.ObjectStateUDT.FieldIdentifier}.{attributeSet.PropertyName}";
            };

            var propertyAttributeSet = CreatePropertyAttributeSets(asUDTCandidate.WrappedCandidate);

            return QualifyBackingField(propertyAttributeSet, set => QualifyPrivateUDTWrappedBackingField(set)).ToList();
        }

        private IEnumerable<PropertyAttributeSet> CreatePropertyAttributeSets(IUserDefinedTypeCandidate candidate)
        {

            if (candidate.TypeDeclarationIsPrivate)
            {
                var allPropertyAttributeSets = new List<PropertyAttributeSet>();
                foreach (var member in candidate.Members)
                {
                    var propertyAttributeSets = CreatePropertyAttributeSets(member);
                    var modifiedSets = QualifyBackingField(propertyAttributeSets, propertyAttributeSet => _backingFieldQualifierFunc(candidate, propertyAttributeSet.BackingField));
                    allPropertyAttributeSets.AddRange(modifiedSets);
                }
                return allPropertyAttributeSets;
            }

            return new List<PropertyAttributeSet>() { CreatePropertyAttributeSet(candidate) };
        }

        private IEnumerable<PropertyAttributeSet> CreatePropertyAttributeSets(IUserDefinedTypeMemberCandidate udtMemberCandidate)
        {
            if (udtMemberCandidate.WrappedCandidate is IUserDefinedTypeCandidate udtCandidate)
            {
                var propertyAttributeSets = CreatePropertyAttributeSets(udtMemberCandidate.WrappedCandidate);

                return udtCandidate.TypeDeclarationIsPrivate
                    ? propertyAttributeSets
                    : QualifyBackingField(propertyAttributeSets, attr => attr.PropertyName);
            }

            return new List<PropertyAttributeSet>() { CreatePropertyAttributeSet(udtMemberCandidate) };
        }

        private IEnumerable<PropertyAttributeSet> CreatePropertyAttributeSets(IEncapsulateFieldCandidate candidate)
        {
            switch (candidate)
            {
                case IUserDefinedTypeCandidate udtCandidate:
                    return CreatePropertyAttributeSets(udtCandidate);
                case IUserDefinedTypeMemberCandidate udtMemberCandidate:
                    return CreatePropertyAttributeSets(udtMemberCandidate);
                default:
                    return new List<PropertyAttributeSet>() { CreatePropertyAttributeSet(candidate) };
            }
        }

        private IEnumerable<PropertyAttributeSet> QualifyBackingField(IEnumerable<PropertyAttributeSet> propertyAttributeSets, Func<PropertyAttributeSet, string> backingFieldQualifier)
        {
            var modifiedSets = new List<PropertyAttributeSet>();
            for (var idx = 0; idx < propertyAttributeSets.Count(); idx++)
            {
                var propertyAttributeSet = propertyAttributeSets.ElementAt(idx);
                propertyAttributeSet.BackingField = backingFieldQualifier(propertyAttributeSet);
                modifiedSets.Add(propertyAttributeSet);
            }
            return modifiedSets;
        }

        private PropertyAttributeSet CreatePropertyAttributeSet(IEncapsulateFieldCandidate candidate)
        {
            return new PropertyAttributeSet()
            {
                PropertyName = candidate.PropertyIdentifier,
                BackingField = candidate.BackingIdentifier,
                AsTypeName = candidate.PropertyAsTypeName,
                RHSParameterIdentifier = Resources.Refactorings.Refactorings.CodeBuilder_DefaultPropertyRHSParam,
                GeneratePropertyLet = !candidate.IsReadOnly && !candidate.Declaration.IsObject && !candidate.Declaration.IsArray,
                GeneratePropertySet = !candidate.IsReadOnly && !candidate.Declaration.IsArray && (candidate.Declaration.IsObject || candidate.Declaration.AsTypeName == Tokens.Variant),
                UsesSetAssignment = candidate.Declaration.IsObject,
                Declaration = candidate.Declaration
            };
        }
    }
}
