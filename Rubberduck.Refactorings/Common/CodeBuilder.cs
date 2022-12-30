using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Common;
using Rubberduck.Resources;
using Rubberduck.SmartIndenter;
using System;
using System.Collections.Generic;
using System.Linq;
using Tokens = Rubberduck.Resources.Tokens;

namespace Rubberduck.Refactorings
{
    public interface ICodeBuilder
    {
        /// <summary>
        /// Returns ModuleBodyElementDeclaration signature with an ImprovedArgument list
        /// </summary>
        string ImprovedFullMemberSignature(ModuleBodyElementDeclaration declaration);

        /// <summary>
        /// Returns a ModuleBodyElementDeclaration block
        /// with an ImprovedArgument List
        /// </summary>
        /// <param name="content">Main body content/logic of the member</param>
        string BuildMemberBlockFromPrototype(ModuleBodyElementDeclaration declaration,
            string content = null,
            Accessibility accessibility = Accessibility.Public,
            string newIdentifier = null);

        /// <summary>
        /// Returns the argument list for the input ModuleBodyElementDeclaration with the following improvements:
        /// 1. Explicitly declares Property Let\Set value parameter as ByVal
        /// 2. Ensures UserDefined Type parameters are declared either explicitly or implicitly as ByRef
        /// </summary>
        string ImprovedArgumentList(ModuleBodyElementDeclaration declaration);

        /// <summary>
        /// Generates a Property Get codeblock based on the prototype declaration 
        /// </summary>
        /// <param name="prototype">DeclarationType with flags: Variable, Constant,
        /// UserDefinedTypeMember, Function, or Property</param>
        /// <param name="content">Member body content.</param>
        /// <param name="parameterIdentifier">Defaults to '<paramref name="propertyIdentifier"/>Value' unless otherwise specified</param>
        bool TryBuildPropertyGetCodeBlock(Declaration prototype,
            string propertyIdentifier,
            out string codeBlock,
            Accessibility accessibility = Accessibility.Public,
            string content = null);

        /// <summary>
        /// Generates a Property Let codeblock based on the prototype declaration 
        /// </summary>
        /// <param name="prototype">DeclarationType with flags: Variable, Constant,
        /// UserDefinedTypeMember, Function, or Property</param>
        /// <param name="content">Member body content.</param>
        /// <param name="parameterIdentifier">Defaults to 'RHS' unless otherwise specified</param>
        bool TryBuildPropertyLetCodeBlock(Declaration prototype,
            string propertyIdentifier,
            out string codeBlock,
            Accessibility accessibility = Accessibility.Public,
            string content = null,
            string parameterIdentifier = null);

        /// <summary>
        /// Generates a Property Set codeblock based on the prototype declaration 
        /// </summary>
        /// <param name="prototype">DeclarationType with flags: Variable, Constant,
        /// UserDefinedTypeMember, Function, or Property</param>
        /// <param name="content">Member body content.</param>
        /// <param name="parameterIdentifier">Defaults to 'RHS' unless otherwise specified</param>
        bool TryBuildPropertySetCodeBlock(Declaration prototype,
            string propertyIdentifier,
            out string codeBlock,
            Accessibility accessibility = Accessibility.Public,
            string content = null,
            string parameterIdentifier = null);

        /// <summary>
        /// Generates a UserDefinedType (UDT) declaration using the prototype declarations for
        /// creating the UserDefinedTypeMember declarations.
        /// </summary>
        /// <remarks>
        /// No validation or conflict analysis is applied to the identifiers.
        /// </remarks>
        /// <param name="memberPrototypes">DeclarationTypes with flags: Variable, Constant, 
        /// UserDefinedTypeMember, Function, or Property</param>
        bool TryBuildUserDefinedTypeDeclaration(string udtIdentifier, 
            IEnumerable<(Declaration Prototype, string UDTMemberIdentifier)> memberPrototypes, 
            out string codeBlock, 
            Accessibility accessibility = Accessibility.Private);

        /// <summary>
        /// Generates a <c>UserDefinedTypeMember</c> declaration expression based on the prototype declaration
        /// </summary>
        /// <remarks>
        /// No validation or conflict analysis is applied to the identifiers.
        /// </remarks>
        /// <param name="prototype">DeclarationType with flags: Variable, Constant, 
        /// UserDefinedTypeMember, Function, or Property</param>
        bool TryBuildUDTMemberDeclaration(Declaration prototype, string identifier, out string codeBlock);

        IIndenter Indenter { get; }
    }

    public class CodeBuilder : ICodeBuilder
    {
        private const string paramSeparator = ", ";

        public CodeBuilder(IIndenter indenter)
        {
            Indenter = indenter;
        }

        public IIndenter Indenter { get; }

        public string BuildMemberBlockFromPrototype(ModuleBodyElementDeclaration declaration, 
            string content = null, 
            Accessibility accessibility = Accessibility.Public, 
            string newIdentifier = null)
        {
            var elements = new List<string>()
            {
                ImprovedFullMemberSignatureInternal(declaration, accessibility, newIdentifier),
                Environment.NewLine,
                string.IsNullOrEmpty(content) ? null : $"{content}{Environment.NewLine}",
                EndStatement(declaration.DeclarationType),
                Environment.NewLine,
            };
            return string.Join(Environment.NewLine, Indenter.Indent(string.Concat(elements)));
        }

        public bool TryBuildPropertyGetCodeBlock(Declaration prototype, 
            string propertyIdentifier, 
            out string codeBlock, 
            Accessibility accessibility = Accessibility.Public, 
            string content = null)

            => TryBuildPropertyBlockFromPrototype(prototype, DeclarationType.PropertyGet, 
                propertyIdentifier, out codeBlock, accessibility, content);

        public bool TryBuildPropertyLetCodeBlock(Declaration prototype,
            string propertyIdentifier, out string codeBlock,
            Accessibility accessibility = Accessibility.Public,
            string content = null, string valueParameterIdentifier = null)
        {
            codeBlock = string.Empty;
            if (IsMutatorPropertyForObjectType(prototype))
            {
                return false;
            }

            return TryBuildPropertyBlockFromPrototype(prototype, DeclarationType.PropertyLet,
                           propertyIdentifier, out codeBlock, accessibility, content, valueParameterIdentifier);
        }

        public bool TryBuildPropertySetCodeBlock(Declaration prototype,
            string propertyIdentifier, out string codeBlock,
            Accessibility accessibility = Accessibility.Public,
            string content = null, string valueParameterIdentifier = null)
        {
            codeBlock = string.Empty;
            if (prototype.IsMutatorProperty())
            {
                var prototypeAsTypeName = AsTypeNameFromMutatorProperty(prototype);
                if (!(prototypeAsTypeName == Tokens.Variant
                        || IsMutatorPropertyForObjectType(prototype)))

                {
                    return false;
                }
            }

            return TryBuildPropertyBlockFromPrototype(prototype, DeclarationType.PropertySet,
                           propertyIdentifier, out codeBlock, accessibility, content, valueParameterIdentifier);
        }

        private bool TryBuildPropertyBlockFromPrototype(Declaration prototype, 
            DeclarationType letSetGetTypeToCreate, string propertyIdentifier, 
            out string codeBlock, Accessibility accessibility, 
            string memberBody = null, string valueParameterIdentifier = null)

        {
            codeBlock = string.Empty;
            if (!IsValidPrototypeDeclarationType(prototype.DeclarationType))
            {
                return false;
            }

            var methodName = $"{TypeToken(letSetGetTypeToCreate)} {propertyIdentifier}";

            var propertyImplementation = memberBody ?? DefaultPropertyImplementation();

            if (letSetGetTypeToCreate.HasFlag(DeclarationType.PropertyGet))
            {
                codeBlock = CreateGetPropertyBlock(prototype, accessibility, methodName, propertyImplementation);
            }
            else
            {
                codeBlock = CreateLetSetPropertyBlock(prototype, letSetGetTypeToCreate,
                    accessibility, methodName, valueParameterIdentifier, memberBody ?? propertyImplementation);
            }

            codeBlock = string.Join(Environment.NewLine, Indenter.Indent(codeBlock));
            return true;
        }

        private static string CreateLetSetPropertyBlock(Declaration prototype, DeclarationType declarationTypeToCreate,
            Accessibility accessibility, string methodName, string valueParameterIdentifier, string memberBody)
        {
            var parameterList = CreateLetSetParameterList(prototype, valueParameterIdentifier);

            var codeBlock = string.Join(
                Environment.NewLine,
                $"{AccessibilityToken(accessibility)} {methodName}({parameterList})",
                memberBody,
                EndStatement(declarationTypeToCreate));

            return codeBlock;
        }

        private static string CreateGetPropertyBlock(Declaration prototype, Accessibility accessibility,
            string methodName, string memberBody)
        {
            var parameters = prototype is IParameterizedDeclaration parameterizedDeclaration
                ? parameterizedDeclaration.Parameters
                    .TakeWhile(p => p != parameterizedDeclaration.Parameters.Last())
                    .Select(GetParameterExpression)
                : Enumerable.Empty<string>();
            
            var parameterList = string.Join(paramSeparator, parameters);

            var asTypeClause = $"{Tokens.As} {PrototypeToPropertyAsTypeName(prototype)}";

            return string.Join(
                Environment.NewLine,
                $"{AccessibilityToken(accessibility)} {methodName}({parameterList}) {asTypeClause}",
                memberBody,
                EndStatement(DeclarationType.PropertyGet));
        }

        public string ImprovedFullMemberSignature(ModuleBodyElementDeclaration declaration)
            => ImprovedFullMemberSignatureInternal(declaration, declaration.Accessibility);

        private string ImprovedFullMemberSignatureInternal(ModuleBodyElementDeclaration declaration,
            Accessibility accessibility, string newIdentifier = null)
        {
            var asTypeName = string.IsNullOrEmpty(declaration.AsTypeName)
                ? string.Empty
                : $" {Tokens.As} {declaration.AsTypeName}";
            
            var elements = new List<string>()
            {
                AccessibilityToken(accessibility),
                $" {TypeToken(declaration.DeclarationType)} ",
                newIdentifier ?? declaration.IdentifierName,
                $"({ImprovedArgumentList(declaration)})",
                asTypeName
            };

            return string.Concat(elements).Trim();
        }

        public string ImprovedArgumentList(ModuleBodyElementDeclaration declaration)
        {
            var arguments = Enumerable.Empty<string>();
            if (declaration is IParameterizedDeclaration parameterizedDeclaration)
            {
                arguments = parameterizedDeclaration.Parameters
                    .OrderBy(parameter => parameter.Selection)
                    .Select(parameter => BuildParameterDeclaration(
                        parameter,
                        parameter.Equals(parameterizedDeclaration.Parameters.LastOrDefault())
                            && declaration.DeclarationType.HasFlag(DeclarationType.Property)
                            && !declaration.DeclarationType.Equals(DeclarationType.PropertyGet)));
            }

            return $"{string.Join(paramSeparator, arguments)}";
        }

        private static string BuildParameterDeclaration(ParameterDeclaration parameter, bool forceExplicitByValAccess)
        {
            var optionalParamType = parameter.IsParamArray
               ? Tokens.ParamArray
               : parameter.IsOptional ? Tokens.Optional : string.Empty;

            var paramMechanism = parameter.IsImplicitByRef
                ? string.Empty
                : parameter.IsByRef ? Tokens.ByRef : Tokens.ByVal;

            if (forceExplicitByValAccess
                && (string.IsNullOrEmpty(paramMechanism) || paramMechanism.Equals(Tokens.ByRef))
                && !parameter.IsUserDefinedType())
            {
                paramMechanism = Tokens.ByVal;
            }

            var name = parameter.IsArray
                ? $"{parameter.IdentifierName}()"
                : parameter.IdentifierName;

            var paramDeclarationElements = new List<string>()
            {
                FormatOptionalElement(optionalParamType),
                FormatOptionalElement(paramMechanism),
                $"{name} ",
                FormatAsTypeName(parameter.AsTypeName),
                FormatDefaultValue(parameter.DefaultValue)
            };

            return string.Concat(paramDeclarationElements).Trim();
        }

        private static string FormatOptionalElement(string element)
            => string.IsNullOrEmpty(element) ? string.Empty : $"{element} ";

        private static string FormatAsTypeName(string AsTypeName) 
            => string.IsNullOrEmpty(AsTypeName) ? string.Empty : $"As {AsTypeName} ";

        private static string FormatDefaultValue(string DefaultValue) 
            => string.IsNullOrEmpty(DefaultValue) ? string.Empty : $"= {DefaultValue}";

        private static Dictionary<DeclarationType, (string TypeToken, string EndStatement)> _declarationTypeTokens
            = new Dictionary<DeclarationType, (string TypeToken, string EndStatement)>()
            {
                [DeclarationType.Function] = (Tokens.Function, $"{Tokens.End} {Tokens.Function}"),
                [DeclarationType.Procedure] = (Tokens.Sub, $"{Tokens.End} {Tokens.Sub}"),
                [DeclarationType.PropertyGet] = ($"{Tokens.Property} {Tokens.Get}", $"{Tokens.End} {Tokens.Property}"),
                [DeclarationType.PropertyLet] = ($"{Tokens.Property} {Tokens.Let}", $"{Tokens.End} {Tokens.Property}"),
                [DeclarationType.PropertySet] = ($"{Tokens.Property} {Tokens.Set}", $"{Tokens.End} {Tokens.Property}"),
            };

        private static string EndStatement(DeclarationType declarationType)
            => _declarationTypeTokens[declarationType].EndStatement;

        private static string TypeToken(DeclarationType declarationType)
            => _declarationTypeTokens[declarationType].TypeToken;

        private static string GetParameterExpression(ParameterDeclaration parameterDeclaration)
        {
            var parameterMechanism = 
                ((VBAParser.ArgContext)parameterDeclaration.Context).BYVAL() == null 
                    ? Tokens.ByRef 
                    : Tokens.ByVal;

            return string.Format("{0} {1} As {2}",
                parameterMechanism,
                parameterDeclaration.IdentifierName,
                parameterDeclaration.AsTypeName);
        }

        public bool TryBuildUserDefinedTypeDeclaration(string udtIdentifier, 
            IEnumerable<(Declaration Prototype, string UDTMemberIdentifier)> memberPrototypes, 
            out string codeBlock, 
            Accessibility accessibility = Accessibility.Private)
        {
            if (udtIdentifier is null
                ||!memberPrototypes.Any()
                || memberPrototypes.Any(p => p.Prototype is null || p.UDTMemberIdentifier is null)
                || memberPrototypes.Any(mp => !IsValidPrototypeDeclarationType(mp.Prototype.DeclarationType)))
            {
                codeBlock = string.Empty;
                return false;
            }

            var blockLines = memberPrototypes
                .Select(m => BuildUDTMemberDeclaration(m.UDTMemberIdentifier, m.Prototype))
                .ToList();

            blockLines.Insert(0, $"{accessibility.TokenString()} {Tokens.Type} {udtIdentifier}");

            blockLines.Add($"{Tokens.End} {Tokens.Type}");

            codeBlock = string.Join(Environment.NewLine, Indenter.Indent(blockLines));

            return true;
        }

        public bool TryBuildUDTMemberDeclaration(Declaration prototype, string udtMemberIdentifier, out string codeBlock)
        {
            codeBlock = string.Empty;

            if (udtMemberIdentifier is null
                || prototype is null
                || !IsValidPrototypeDeclarationType(prototype.DeclarationType))
            {
                return false;
            }

            codeBlock = BuildUDTMemberDeclaration(udtMemberIdentifier, prototype);
            return true;
        }

        private static string BuildUDTMemberDeclaration(string udtMemberIdentifier, Declaration prototype)
        {
            var asTypeName = prototype.AsTypeName;

            if (prototype.IsArray)
            {
                var identifierExpression = prototype.Context.TryGetChildContext<VBAParser.ArrayDimContext>(out var ctxt)
                    ? $"{udtMemberIdentifier}{ctxt.GetText()}"
                    : $"{udtMemberIdentifier}()";   // This should never happen.
                
                return $"{identifierExpression} {Tokens.As} {asTypeName}";
            }

            if (prototype.IsMutatorProperty())
            {
                asTypeName = AsTypeNameFromMutatorProperty(prototype);
            }

            return $"{udtMemberIdentifier} {Tokens.As} {asTypeName}";
        }

        private static string PrototypeToPropertyAsTypeName(Declaration prototype)
        {
            //TODO: Improve generated Array properties
            //Add logic to conditionally return Arrays or Variant depending on Office version.
            //Ability to return an Array from a Function or 
            //Property was added in Office 2000 http://www.cpearson.com/excel/passingandreturningarrays.htm
            
            if (prototype.IsArray)
            {
                return Tokens.Variant;
            }

            if (prototype.IsMutatorProperty())
            {
                return AsTypeNameFromMutatorProperty(prototype);
            }

            return prototype.IsEnumField()
                ? EnumerationToPropertyAsTypeName(prototype)
                : prototype.AsTypeName;
        }

        private static string AsTypeNameFromMutatorProperty(Declaration prototype)
        {
            var paramDeclaration = prototype as IParameterizedDeclaration;
            return paramDeclaration.Parameters.Last().AsTypeName;
        }

        private static string CreateLetSetParameterList(Declaration prototype, string valueParameterIdentifier = null)
        {
            if (prototype.IsMutatorProperty())
            {
                var parameterizedDeclaration = prototype as IParameterizedDeclaration;
                return string.Join(paramSeparator, parameterizedDeclaration.Parameters.Select(GetParameterExpression));
            }

            var paramMechanism = prototype.IsUserDefinedType() ? Tokens.ByRef : Tokens.ByVal;

            var asTypeClause = $"{Tokens.As} {PrototypeToPropertyAsTypeName(prototype)}";

            var valueParameterName = valueParameterIdentifier
                ?? Resources.Refactorings.Refactorings.CodeBuilder_DefaultPropertyRHSParam;

            var parameters = prototype is IParameterizedDeclaration pDec
                ? pDec.Parameters.Select(GetParameterExpression).ToList() //Property Get prototype
                : new List<string>(); //Variable, UDT Member, Function prototypes

            var valueParameterExpression = $"{paramMechanism} {valueParameterName} {asTypeClause}";
            parameters.Add(valueParameterExpression);

            return string.Join(paramSeparator, parameters);
        }

        private static bool IsMutatorPropertyForObjectType(Declaration prototype) 
            => prototype.IsMutatorProperty() 
                && prototype is IParameterizedDeclaration pd && pd.Parameters.Last().IsObject;

        private static string AccessibilityToken(Accessibility accessibility)
            => accessibility.Equals(Accessibility.Implicit)
                ? Tokens.Public
                : $"{accessibility}";

        private static bool IsValidPrototypeDeclarationType(DeclarationType declarationType)
        {
            return declarationType.HasFlag(DeclarationType.Variable)
                || declarationType.HasFlag(DeclarationType.UserDefinedTypeMember)
                || declarationType.HasFlag(DeclarationType.Constant)
                || declarationType.HasFlag(DeclarationType.Function)
                || declarationType.HasFlag(DeclarationType.Property);
        }

        private static string EnumerationToPropertyAsTypeName(Declaration enumeration) 
            => enumeration.AsTypeDeclaration.HasPrivateAccessibility()
                ? Tokens.Long
                : enumeration.AsTypeName;

        private string DefaultPropertyImplementation()
            => $"    {Tokens.Err}.Raise 5 " +
                $"{Resources.Refactorings.Refactorings.CodeBuilder_DefaultPropertyImplementation}";

    }
}
