using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldCodeBuilder
    {
        (string Get, string Let, string Set) BuildPropertyBlocks(PropertyAttributeSet propertyAttributeSet);
        string BuildUserDefinedTypeDeclaration(IObjectStateUDT objectStateUDT, IEnumerable<IEncapsulateFieldCandidate> candidates);
        string BuildObjectStateFieldDeclaration(IObjectStateUDT objectStateUDT);
        string BuildFieldDeclaration(Declaration target, string identifier);
    }

    /// <summary>
    /// EncapsulateFieldCodeBuilder wraps an ICodeBuilder instance to extend it for the 
    /// specific needs of an EncapsulateField refactoring action.
    /// </summary>
    public class EncapsulateFieldCodeBuilder : IEncapsulateFieldCodeBuilder
    {
        private readonly ICodeBuilder _codeBuilder;

        public EncapsulateFieldCodeBuilder(ICodeBuilder codeBuilder)
        {
            _codeBuilder = codeBuilder;
        }

        public (string Get, string Let, string Set) BuildPropertyBlocks(PropertyAttributeSet propertyAttributes)
        {
            if (!(propertyAttributes.Declaration.DeclarationType.HasFlag(DeclarationType.Variable)
                || propertyAttributes.Declaration.DeclarationType.HasFlag(DeclarationType.UserDefinedTypeMember)))
            {
                throw new ArgumentException("Invalid prototype DeclarationType", nameof(propertyAttributes));
            }

            (string Get, string Let, string Set) blocks = (string.Empty, string.Empty, string.Empty);

            var mutatorBody = $"{propertyAttributes.BackingField} = {propertyAttributes.RHSParameterIdentifier}";

            if (propertyAttributes.GeneratePropertyLet)
            {
                _codeBuilder.TryBuildPropertyLetCodeBlock(propertyAttributes.Declaration, propertyAttributes.PropertyName, out blocks.Let, content: mutatorBody);
            }

            if (propertyAttributes.GeneratePropertySet)
            {
                _codeBuilder.TryBuildPropertySetCodeBlock(propertyAttributes.Declaration, propertyAttributes.PropertyName, out blocks.Set, content: $"{Tokens.Set} {mutatorBody}");
            }

            var propertyGetBody = propertyAttributes.UsesSetAssignment
                ? $"{Tokens.Set} {propertyAttributes.PropertyName} = {propertyAttributes.BackingField}"
                : $"{propertyAttributes.PropertyName} = {propertyAttributes.BackingField}";

            if (propertyAttributes.AsTypeName.Equals(Tokens.Variant) && !propertyAttributes.Declaration.IsArray)
            {
                propertyGetBody = string.Join(Environment.NewLine,
                    $"{Tokens.If} IsObject({propertyAttributes.BackingField}) {Tokens.Then}",
                    $"{Tokens.Set} {propertyAttributes.PropertyName} = {propertyAttributes.BackingField}",
                    Tokens.Else,
                    $"{propertyAttributes.PropertyName} = {propertyAttributes.BackingField}",
                    $"{Tokens.End} {Tokens.If}");
            }

            _codeBuilder.TryBuildPropertyGetCodeBlock(propertyAttributes.Declaration, propertyAttributes.PropertyName, out blocks.Get, content: propertyGetBody);

            return (blocks.Get, blocks.Let, blocks.Set);
        }

        public string BuildUserDefinedTypeDeclaration(IObjectStateUDT objectStateUDT, IEnumerable<IEncapsulateFieldCandidate> candidates)
        {
            var newUDTMembers = candidates.Where(c => c.EncapsulateFlag)
                .Select(m => (m.Declaration, m.BackingIdentifier));

            if (_codeBuilder.TryBuildUserDefinedTypeDeclaration(objectStateUDT.AsTypeName, newUDTMembers, out var declaration))
            {
                return declaration;
            }

            return string.Empty;
        }

        public string BuildObjectStateFieldDeclaration(IObjectStateUDT objectStateUDT) 
            => $"{Accessibility.Private} {objectStateUDT.IdentifierName} {Tokens.As} {objectStateUDT.AsTypeName}";

        public string BuildFieldDeclaration(Declaration target, string identifier)
        {
            var identifierExpressionSansVisibility = target.Context.GetText().Replace(target.IdentifierName, identifier);
            return target.IsTypeSpecified
                ? $"{Tokens.Private} {identifierExpressionSansVisibility}"
                : $"{Tokens.Private} {identifierExpressionSansVisibility} {Tokens.As} {target.AsTypeName}";
        }
    }
}
