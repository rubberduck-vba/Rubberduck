using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public static class EncapsulateFieldExtensions
    {
        public static bool IsVariable(this Declaration declaration)
            => declaration.DeclarationType.HasFlag(DeclarationType.Variable);

        public static bool IsMemberVariable(this Declaration declaration)
            => declaration.IsVariable() && !declaration.ParentDeclaration.IsMember();

        public static bool IsLocalVariable(this Declaration declaration)
            => declaration.IsVariable() && declaration.ParentDeclaration.IsMember();

        public static bool IsLocalConstant(this Declaration declaration)
            => declaration.IsConstant() && declaration.ParentDeclaration.IsMember();

        public static bool HasPrivateAccessibility(this Declaration declaration)
            => declaration.Accessibility.Equals(Accessibility.Private);

        public static bool IsMember(this Declaration declaration)
            => declaration.DeclarationType.HasFlag(DeclarationType.Member);

        public static bool IsConstant(this Declaration declaration)
            => declaration.DeclarationType.HasFlag(DeclarationType.Constant);

        public static bool IsUserDefinedTypeField(this Declaration declaration)
            => declaration.IsMemberVariable() && (declaration.AsTypeDeclaration?.DeclarationType.Equals(DeclarationType.UserDefinedType) ?? false);

        public static bool IsEnumField(this Declaration declaration)
            => declaration.IsMemberVariable() && (declaration.AsTypeDeclaration?.DeclarationType.Equals(DeclarationType.Enumeration) ?? false);

        public static bool IsDeclaredInList(this Declaration declaration)
        {
            return declaration.Context.TryGetAncestor<VBAParser.VariableListStmtContext>(out var varList)
                            && varList.ChildCount > 1;
        }

        public static IEnumerable<IdentifierReference> AllReferences(this IEnumerable<Declaration> declarations)
        {
            return from dec in declarations
                   from reference in dec.References
                   select reference;
        }

        public static string Capitalize(this string input)
            => $"{char.ToUpperInvariant(input[0]) + input.Substring(1, input.Length - 1)}";

        public static string UnCapitalize(this string input)
            => $"{char.ToLowerInvariant(input[0]) + input.Substring(1, input.Length - 1)}";

        public static bool IsEquivalentVBAIdentifierTo(this string lhs, string identifier)
            => lhs.Equals(identifier, StringComparison.InvariantCultureIgnoreCase);

        public static string GetText(this IModuleRewriter rewriter, int maxConsecutiveNewLines)
        {
            var result = rewriter.GetText();
            var target = string.Join(string.Empty, Enumerable.Repeat(Environment.NewLine, maxConsecutiveNewLines).ToList());
            var replacement = string.Join(string.Empty, Enumerable.Repeat(Environment.NewLine, maxConsecutiveNewLines - 1).ToList());
            for (var counter = 1; counter < 10 && result.Contains(target); counter++)
            {
                result = result.Replace(target, replacement);
            }
            return result;
        }

        public static string IncrementEncapsulationIdentifier(this string identifier)
        {
            var fragments = identifier.Split('_');
            if (fragments.Length == 1) { return $"{identifier}_1"; }

            var lastFragment = fragments[fragments.Length - 1];
            if (long.TryParse(lastFragment, out var number))
            {
                fragments[fragments.Length - 1] = (number + 1).ToString();

                return string.Join("_", fragments);
            }
            return $"{identifier}_1"; ;
        }

        public static void InsertAtEndOfFile(this IModuleRewriter rewriter, string content)
        {
            if (content == string.Empty) { return; }

            rewriter.InsertBefore(rewriter.TokenStream.Size - 1, content);
        }

        public static void MakeImplicitDeclarationTypeExplicit(this IModuleRewriter rewriter, Declaration element)
        {
            if (!element.Context.TryGetChildContext<VBAParser.AsTypeClauseContext>(out _))
            {
                rewriter.InsertAfter(element.Context.Stop.TokenIndex, $" {Tokens.As} {element.AsTypeName}");
            }
        }

        public static void Rename(this IModuleRewriter rewriter, Declaration target, string newName)
        {
            if (target.Context is IIdentifierContext context)
            {
                rewriter.Replace(context.IdentifierTokens, newName);
            }
        }

        public static void SetVariableVisiblity(this IModuleRewriter rewriter, Declaration element, string visibility)
        {
            if (!element.IsVariable()) { throw new ArgumentException(); }

            var variableStmtContext = element.Context.GetAncestor<VBAParser.VariableStmtContext>();
            var visibilityContext = variableStmtContext.GetChild<VBAParser.VisibilityContext>();

            if (visibilityContext != null)
            {
                rewriter.Replace(visibilityContext, visibility);
                return;
            }
            rewriter.InsertBefore(element.Context.Start.TokenIndex, $"{visibility} ");
        }
    }

}
