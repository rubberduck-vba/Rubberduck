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
    public static class StringExtensions
    {
        public static bool IsEquivalentVBAIdentifierTo(this string lhs, string identifier)
                => lhs.Equals(identifier, StringComparison.InvariantCultureIgnoreCase);


        public static string IncrementIdentifier2(this string identifier)
        {
            var fragments = identifier.Split('x');
            if (fragments.Length == 1) { return $"{identifier}x1"; }

            var lastFragment = fragments[fragments.Length - 1];
            if (long.TryParse(lastFragment, out var number))
            {
                fragments[fragments.Length - 1] = (number + 1).ToString();

                return string.Join("x", fragments);
            }
            return $"{identifier}x1"; ;
        }

        public static string IncrementIdentifier(this string identifier)
        {
            var numeric = string.Join(string.Empty, identifier.Reverse().TakeWhile(c => char.IsDigit(c)).Reverse());
            if (!int.TryParse(numeric, out var currentNum))
            {
                currentNum = 0;
            }
            var identifierSansNumericSuffix = identifier.Substring(0, identifier.Length - numeric.Length);
            return $"{identifierSansNumericSuffix}{++currentNum}";
        }
    }

    public static class IParameterizedDeclarationExtensions
    {
        public static string BuildFullyDefinedArgumentList(this IParameterizedDeclaration pDeclaration)
        {
            string AsFullyDefinedParameter(ParameterDeclaration p)
            {
                var access = ((VBAParser.ArgContext)p.Context).BYVAL() != null
                    ? Tokens.ByVal
                    : Tokens.ByRef;

                if (p.AsTypeDeclaration?.DeclarationType.Equals(DeclarationType.UserDefinedType) ?? false
                    || (p == pDeclaration.Parameters.Last()
                            && (p.ParentDeclaration.DeclarationType.Equals(DeclarationType.PropertyLet)
                                    || p.ParentDeclaration.DeclarationType.Equals(DeclarationType.PropertySet))))
                {
                    access = Tokens.ByVal;
                }

                return $"{access} {p.IdentifierName} {Tokens.As} {p.AsTypeName}";
            }

            var memberParams = pDeclaration.Parameters.ToList()
                .OrderBy(o => o.Selection.StartLine)
                .ThenBy(t => t.Selection.StartColumn)
                .Select(p => AsFullyDefinedParameter(p));

            var improvedArgList = $"({string.Join(", ", memberParams)})";
            return improvedArgList;
        }
    }
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

        //public static bool ContainsParentScopeToAllReferences(this IEnumerable<Declaration> containing, IEnumerable<Declaration> declarations)
        //{
        //    return ContainsParentScopeToAllReferences(containing, declarations.AllReferences());
        //}

        public static bool ContainsParentScopesForAllReferences(this IEnumerable<Declaration> containing, IEnumerable<IdentifierReference> references)
        {
            return references.All(rf => containing.Contains(rf.ParentScoping));
        }

        public static bool ContainsParentScopeForAnyReference(this IEnumerable<Declaration> containing, IEnumerable<IdentifierReference> references)
        {
            return references.Any(rf => containing.Contains(rf.ParentScoping));
        }
    }
}
