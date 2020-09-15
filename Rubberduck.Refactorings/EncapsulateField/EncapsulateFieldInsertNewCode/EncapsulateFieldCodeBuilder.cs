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

    public class EncapsulateFieldCodeBuilder : IEncapsulateFieldCodeBuilder
    {
        private const string FourSpaces = "    ";
        private static string _doubleSpace = $"{Environment.NewLine}{Environment.NewLine}";

        private readonly ICodeBuilder _codeBuilder;

        public EncapsulateFieldCodeBuilder(ICodeBuilder codeBuilder)
        {
            _codeBuilder = codeBuilder;
        }

        public (string Get, string Let, string Set) BuildPropertyBlocks(PropertyAttributeSet propertyAttributes)
        {
            string propertyLet = null;
            string propertySet = null;

            if (propertyAttributes.GeneratePropertyLet)
            {
                var letterContent = $"{FourSpaces}{propertyAttributes.BackingField} = {propertyAttributes.RHSParameterIdentifier}";
                if (!_codeBuilder.TryBuildPropertyLetCodeBlock(propertyAttributes.Declaration, propertyAttributes.PropertyName, out propertyLet, content: letterContent))
                {
                    throw new ArgumentException();
                }
            }

            if (propertyAttributes.GeneratePropertySet)
            {
                var setterContent = $"{FourSpaces}{Tokens.Set} {propertyAttributes.BackingField} = {propertyAttributes.RHSParameterIdentifier}";
                if (!_codeBuilder.TryBuildPropertySetCodeBlock(propertyAttributes.Declaration, propertyAttributes.PropertyName, out propertySet, content: setterContent))
                {
                    throw new ArgumentException();
                }
            }

            var getterContent = $"{propertyAttributes.PropertyName} = {propertyAttributes.BackingField}";
            if (propertyAttributes.UsesSetAssignment)
            {
                getterContent = $"{Tokens.Set} {getterContent}";
            }

            if (propertyAttributes.AsTypeName.Equals(Tokens.Variant) && !propertyAttributes.Declaration.IsArray)
            {
                getterContent = string.Join(Environment.NewLine,
                    $"{Tokens.If} IsObject({propertyAttributes.BackingField}) {Tokens.Then}",
                    $"{FourSpaces}{Tokens.Set} {propertyAttributes.PropertyName} = {propertyAttributes.BackingField}",
                    Tokens.Else,
                    $"{FourSpaces}{propertyAttributes.PropertyName} = {propertyAttributes.BackingField}",
                    $"{Tokens.End} {Tokens.If}",
                    Environment.NewLine);
            }

            if (!_codeBuilder.TryBuildPropertyGetCodeBlock(propertyAttributes.Declaration, propertyAttributes.PropertyName, out var propertyGet, content: $"{FourSpaces}{getterContent}"))
            {
                throw new ArgumentException();
            }

            return (propertyGet, propertyLet, propertySet);
        }

        public string BuildUserDefinedTypeDeclaration(IObjectStateUDT objectStateUDT, IEnumerable<IEncapsulateFieldCandidate> candidates)
        {
            var selected = candidates.Where(c => c.EncapsulateFlag);

            var newUDTMembers = selected
                .Select(m => (m.Declaration, m.BackingIdentifier));

            _codeBuilder.TryBuildUserDefinedTypeDeclaration(objectStateUDT.AsTypeName, newUDTMembers, out var declaration);

            return declaration ?? string.Empty;
        }

        public string BuildObjectStateFieldDeclaration(IObjectStateUDT objectStateUDT)
        {
            return $"{Accessibility.Private} {objectStateUDT.IdentifierName} {Tokens.As} {objectStateUDT.AsTypeName}";
        }

        public string BuildFieldDeclaration(Declaration target, string identifier)
        {
            var identifierExpressionSansVisibility = target.Context.GetText().Replace(target.IdentifierName, identifier);
            return target.IsTypeSpecified
                ? $"{Tokens.Private} {identifierExpressionSansVisibility}"
                : $"{Tokens.Private} {identifierExpressionSansVisibility} {Tokens.As} {target.AsTypeName}";
        }
    }
}
