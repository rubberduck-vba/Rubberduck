using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Adjusts a Function procedure to be a Sub procedure. This operation may result in broken code.
    /// </summary>
    /// <inspections>
    /// <inspection name="NonReturningFunctionInspection" />
    /// <inspection name="FunctionReturnValueAlwaysDiscardedInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="false" module="true" project="false" all="false" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim value As Long
    ///     GetValue value
    ///     Debug.Print value
    /// End Sub
    /// 
    /// Private Function GetValue(ByRef value As Long)
    ///     value = 42
    /// End Function
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim value As Long
    ///     GetValue value
    ///     Debug.Print value
    /// End Sub
    /// 
    /// Private Sub GetValue(ByRef value As Long)
    ///     value = 42
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class ConvertToProcedureQuickFix : QuickFixBase
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public ConvertToProcedureQuickFix(IDeclarationFinderProvider declarationFinderProvider)
            : base(typeof(NonReturningFunctionInspection), typeof(FunctionReturnValueAlwaysDiscardedInspection))
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            if (!(result.Target is ModuleBodyElementDeclaration moduleBodyElementDeclaration))
            {
                return;
            }

            if (moduleBodyElementDeclaration.IsInterfaceMember)
            {
                var implementations = _declarationFinderProvider
                    .DeclarationFinder
                    .FindInterfaceImplementationMembers(moduleBodyElementDeclaration);
                foreach (var implementation in implementations)
                {
                    ConvertMember(implementation, rewriteSession);
                }
            }

            ConvertMember(moduleBodyElementDeclaration, rewriteSession);
        }

        private void ConvertMember(ModuleBodyElementDeclaration member, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(member.QualifiedModuleName);

            switch (member.Context)
            {
                case VBAParser.FunctionStmtContext functionContext:
                    ConvertFunction(member, functionContext, rewriter);
                    break;
                case VBAParser.PropertyGetStmtContext propertyGetContext:
                    ConvertPropertyGet(member, propertyGetContext, rewriter);
                    break;
            }
        }

        private void ConvertFunction(ModuleBodyElementDeclaration member, VBAParser.FunctionStmtContext functionContext, IModuleRewriter rewriter)
        {
            RemoveAsTypeDeclaration(functionContext, rewriter);
            RemoveTypeHint(member, functionContext, rewriter);

            ConvertFunctionDeclaration(functionContext, rewriter);
            ConvertExitFunctionStatements(functionContext, rewriter);

            RemoveReturnStatements(member, rewriter);
        }

        private static void RemoveAsTypeDeclaration(ParserRuleContext functionContext, IModuleRewriter rewriter)
        {
            var asTypeContext = functionContext.GetChild<VBAParser.AsTypeClauseContext>();
            if (asTypeContext != null)
            {
                rewriter.Remove(asTypeContext);
                rewriter.Remove(
                    functionContext.children.ElementAt(functionContext.children.IndexOf(asTypeContext) -
                                                       1) as ParserRuleContext);
            }
        }

        private static void RemoveTypeHint(ModuleBodyElementDeclaration member, ParserRuleContext functionContext, IModuleRewriter rewriter)
        {
            if (member.TypeHint != null)
            {
                rewriter.Remove(functionContext.GetDescendent<VBAParser.TypeHintContext>());
            }
        }

        private void RemoveReturnStatements(ModuleBodyElementDeclaration member, IModuleRewriter rewriter)
        {
            foreach (var returnStatement in GetReturnStatements(member))
            {
                rewriter.Remove(returnStatement);
            }
        }

        private static void ConvertFunctionDeclaration(VBAParser.FunctionStmtContext functionContext, IModuleRewriter rewriter)
        {
            rewriter.Replace(functionContext.FUNCTION(), Tokens.Sub);
            rewriter.Replace(functionContext.END_FUNCTION(), "End Sub");
        }

        private static void ConvertExitFunctionStatements(VBAParser.FunctionStmtContext context, IModuleRewriter rewriter)
        {
            var exitStatements = context.GetDescendents<VBAParser.ExitStmtContext>();
            foreach (var exitStatement in exitStatements)
            {
                if (exitStatement.EXIT_FUNCTION() != null)
                {
                    rewriter.Replace(exitStatement, $"{Tokens.Exit} {Tokens.Sub}");
                }
            }
        }

        private void ConvertPropertyGet(ModuleBodyElementDeclaration member, VBAParser.PropertyGetStmtContext propertyGetContext, IModuleRewriter rewriter)
        {
            RemoveAsTypeDeclaration(propertyGetContext, rewriter);
            RemoveTypeHint(member, propertyGetContext, rewriter);

            ConvertPropertyGetDeclaration(propertyGetContext, rewriter);
            ConvertExitPropertyStatements(propertyGetContext, rewriter);

            RemoveReturnStatements(member, rewriter);
        }

        private static void ConvertPropertyGetDeclaration(VBAParser.PropertyGetStmtContext propertyGetContext, IModuleRewriter rewriter)
        {
            rewriter.Replace(propertyGetContext.PROPERTY_GET(), Tokens.Sub);
            rewriter.Replace(propertyGetContext.END_PROPERTY(), "End Sub");
        }

        private static void ConvertExitPropertyStatements(VBAParser.PropertyGetStmtContext context, IModuleRewriter rewriter)
        {
            var exitStatements = context.GetDescendents<VBAParser.ExitStmtContext>();
            foreach (var exitStatement in exitStatements)
            {
                if (exitStatement.EXIT_PROPERTY() != null)
                {
                    rewriter.Replace(exitStatement, $"{Tokens.Exit} {Tokens.Sub}");
                }
            }
        }

        private IEnumerable<ParserRuleContext> GetReturnStatements(Declaration declaration)
        {
            return declaration.References
                .Where(usage => IsReturnStatement(declaration, usage))
                .Select(usage => usage.Context.Parent)
                .Cast<ParserRuleContext>();
        }

        private bool IsReturnStatement(Declaration declaration, IdentifierReference assignment)
        {
            return assignment.ParentScoping.Equals(declaration) && assignment.Declaration.Equals(declaration);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.ConvertFunctionToProcedureQuickFix;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => false;
        public override bool CanFixAll => false;
    }
}
