using System;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.InternalApi.Extensions;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Changes 16-bit (max value 32,767) Integer declarations to use 32-bit (max value 2,147,483,647‬) Long integer type instead.
    /// </summary>
    /// <inspections>
    /// <inspection name="IntegerDataTypeInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="true" module="true" project="true" all="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim row As Integer
    ///     row = Sheet1.Range("A" & Sheet1.Rows.Count).End(xlUp).Row
    ///     Debug.Print row
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim row As Long
    ///     row = Sheet1.Range("A" & Sheet1.Rows.Count).End(xlUp).Row
    ///     Debug.Print row
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class ChangeIntegerToLongQuickFix : QuickFixBase
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public ChangeIntegerToLongQuickFix(IDeclarationFinderProvider declarationFinderProvider)
            : base(typeof(IntegerDataTypeInspection))
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.Target.QualifiedModuleName);

            if (result.Target.HasTypeHint)
            {
                ReplaceTypeHint(result.Context, rewriter);
            }
            else
            {
                switch (result.Target.DeclarationType)
                {
                    case DeclarationType.Variable:
                        var variableContext = (VBAParser.VariableSubStmtContext)result.Context;
                        rewriter.Replace(variableContext.asTypeClause().type(), Tokens.Long);
                        break;
                    case DeclarationType.Constant:
                        var constantContext = (VBAParser.ConstSubStmtContext)result.Context;
                        rewriter.Replace(constantContext.asTypeClause().type(), Tokens.Long);
                        break;
                    case DeclarationType.Parameter:
                        var parameterContext = (VBAParser.ArgContext)result.Context;
                        rewriter.Replace(parameterContext.asTypeClause().type(), Tokens.Long);
                        break;
                    case DeclarationType.Function:
                        var functionContext = (VBAParser.FunctionStmtContext)result.Context;
                        rewriter.Replace(functionContext.asTypeClause().type(), Tokens.Long);
                        break;
                    case DeclarationType.PropertyGet:
                        var propertyContext = (VBAParser.PropertyGetStmtContext)result.Context;
                        rewriter.Replace(propertyContext.asTypeClause().type(), Tokens.Long);
                        break;
                    case DeclarationType.UserDefinedTypeMember:
                        var userDefinedTypeMemberContext = (VBAParser.UdtMemberContext)result.Context;
                        rewriter.Replace(
                            userDefinedTypeMemberContext.reservedNameMemberDeclaration() == null
                                ? userDefinedTypeMemberContext.untypedNameMemberDeclaration().optionalArrayClause().asTypeClause().type()
                                : userDefinedTypeMemberContext.reservedNameMemberDeclaration().asTypeClause().type(),
                            Tokens.Long);
                        break;
                }
            }

            var interfaceMembers = _declarationFinderProvider.DeclarationFinder.FindAllInterfaceMembers().ToArray();

            ParserRuleContext matchingInterfaceMemberContext;

            switch (result.Target.DeclarationType)
            {
                case DeclarationType.Parameter:
                    matchingInterfaceMemberContext = interfaceMembers.Select(member => member.Context).FirstOrDefault(c => c == result.Context.Parent.Parent);

                    if (matchingInterfaceMemberContext != null)
                    {
                        var interfaceParameterIndex = GetParameterIndex((VBAParser.ArgContext)result.Context);

                        var implementationMembers =
                            _declarationFinderProvider.DeclarationFinder.FindInterfaceImplementationMembers(interfaceMembers.First(
                                member => member.Context == matchingInterfaceMemberContext)).ToHashSet();

                        var parameterDeclarations =
                            _declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.Parameter)
                                .Where(p => implementationMembers.Contains(p.ParentDeclaration))
                                .Cast<ParameterDeclaration>()
                                .ToArray();

                        foreach (var parameter in parameterDeclarations)
                        {
                            var parameterContext = (VBAParser.ArgContext)parameter.Context;
                            var parameterIndex = GetParameterIndex(parameterContext);

                            if (parameterIndex == interfaceParameterIndex)
                            {
                                var parameterRewriter = rewriteSession.CheckOutModuleRewriter(parameter.QualifiedModuleName);

                                if (parameter.HasTypeHint)
                                {
                                    ReplaceTypeHint(parameterContext, parameterRewriter);
                                }
                                else
                                {
                                    parameterRewriter.Replace(parameterContext.asTypeClause().type(), Tokens.Long);
                                }
                            }
                        }
                    }
                    break;
                case DeclarationType.Function:
                    matchingInterfaceMemberContext = interfaceMembers.Select(member => member.Context).FirstOrDefault(c => c == result.Context);

                    if (matchingInterfaceMemberContext != null)
                    {
                        var functionDeclarations =
                            _declarationFinderProvider.DeclarationFinder.FindInterfaceImplementationMembers(
                                    interfaceMembers.First(member => member.Context == matchingInterfaceMemberContext))
                                .Cast<FunctionDeclaration>()
                                .ToHashSet();

                        foreach (var function in functionDeclarations)
                        {
                            var functionRewriter = rewriteSession.CheckOutModuleRewriter(function.QualifiedModuleName);

                            if (function.HasTypeHint)
                            {
                                ReplaceTypeHint(function.Context, functionRewriter);
                            }
                            else
                            {
                                var functionContext = (VBAParser.FunctionStmtContext)function.Context;
                                functionRewriter.Replace(functionContext.asTypeClause().type(), Tokens.Long);
                            }
                        }
                    }
                    break;
                case DeclarationType.PropertyGet:
                    matchingInterfaceMemberContext = interfaceMembers.Select(member => member.Context).FirstOrDefault(c => c == result.Context);

                    if (matchingInterfaceMemberContext != null)
                    {
                        var propertyGetDeclarations =
                            _declarationFinderProvider.DeclarationFinder.FindInterfaceImplementationMembers(
                                    interfaceMembers.First(member => member.Context == matchingInterfaceMemberContext))
                                .Cast<PropertyGetDeclaration>()
                                .ToHashSet();

                        foreach (var propertyGet in propertyGetDeclarations)
                        {
                            var propertyGetRewriter = rewriteSession.CheckOutModuleRewriter(propertyGet.QualifiedModuleName);

                            if (propertyGet.HasTypeHint)
                            {
                                ReplaceTypeHint(propertyGet.Context, propertyGetRewriter);
                            }
                            else
                            {
                                var propertyGetContext = (VBAParser.PropertyGetStmtContext)propertyGet.Context;
                                propertyGetRewriter.Replace(propertyGetContext.asTypeClause().type(), Tokens.Long);
                            }
                        }
                    }
                    break;
            }
        }

        private static int GetParameterIndex(VBAParser.ArgContext context)
        {
            return Array.IndexOf(((VBAParser.ArgListContext)context.Parent).arg().ToArray(), context);
        }

        private static void ReplaceTypeHint(RuleContext context, IModuleRewriter rewriter)
        {
            var typeHintContext = ((ParserRuleContext)context).GetDescendent<VBAParser.TypeHintContext>();
            rewriter.Replace(typeHintContext, "&");
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.IntegerDataTypeQuickFix;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;
    }
}
