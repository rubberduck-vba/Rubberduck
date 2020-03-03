using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.MoveMember.Extensions
{

    public static class DeclarationExtensions
    {
        public static bool IsVariable(this Declaration declaration)
            => declaration.DeclarationType.HasFlag(DeclarationType.Variable);

        public static bool IsField(this Declaration declaration)
            => declaration.IsVariable() && !declaration.IsLocalVariable();

        public static bool IsLocalVariable(this Declaration declaration)
            => declaration.IsVariable() && declaration.ParentDeclaration.IsMember();

        public static bool IsModuleConstant(this Declaration declaration)
            => declaration.IsConstant() && !declaration.IsLocalConstant();

        public static bool IsConstant(this Declaration declaration)
            => declaration.DeclarationType.HasFlag(DeclarationType.Constant);

        public static bool IsLocalConstant(this Declaration declaration)
            => declaration.IsConstant() && declaration.ParentDeclaration.IsMember();

        public static bool HasPrivateAccessibility(this Declaration declaration)
            => declaration.Accessibility.Equals(Accessibility.Private);

        public static bool IsMember(this Declaration declaration)
            => declaration.DeclarationType.HasFlag(DeclarationType.Member);

        public static IEnumerable<IdentifierReference> AllReferences(this IEnumerable<Declaration> declarations)
        {
            return from dec in declarations
                   from reference in dec.References
                   select reference;
        }

        public static string FullyDefinedSignature(this ModuleBodyElementDeclaration declaration)
        {
            var memberType = string.Empty;
            switch (declaration.Context)
            {
                case VBAParser.SubStmtContext _:
                    memberType = Tokens.Sub;
                    break;
                case VBAParser.FunctionStmtContext _:
                    memberType = Tokens.Function;
                    break;
                case VBAParser.PropertyGetStmtContext _:
                    memberType = $"{Tokens.Property} {Tokens.Get}";
                    break;
                case VBAParser.PropertyLetStmtContext _:
                    memberType = $"{Tokens.Property} {Tokens.Let}";
                    break;
                case VBAParser.PropertySetStmtContext _:
                    memberType = $"{Tokens.Property} {Tokens.Set}";
                    break;
                default:
                    throw new ArgumentException();
            }

            var accessibilityToken = declaration.Accessibility.Equals(Accessibility.Implicit)
                ? Tokens.Public
                : $"{declaration.Accessibility.ToString()}";

            var signature = $"{memberType} {declaration.IdentifierName}()";
            if (declaration is IParameterizedDeclaration parameterizedDeclaration)
            {
                signature = signature.Replace("()", parameterizedDeclaration.BuildFullyDefinedArgumentList());
            }

            var fullSignature = declaration.AsTypeName == null ?
                $"{accessibilityToken} {signature}"
                : $"{accessibilityToken} {signature} As {declaration.AsTypeName}";

            return fullSignature;
        }

        public static bool ContainsParentScopesForAll(this IEnumerable<Declaration> containing, IEnumerable<IdentifierReference> references) 
            => references.All(rf => containing.Contains(rf.ParentScoping));

        public static bool ContainsParentScopeForAnyReference(this IEnumerable<Declaration> containing, IEnumerable<IdentifierReference> references) 
            => references.Any(rf => containing.Contains(rf.ParentScoping));
    }
}
