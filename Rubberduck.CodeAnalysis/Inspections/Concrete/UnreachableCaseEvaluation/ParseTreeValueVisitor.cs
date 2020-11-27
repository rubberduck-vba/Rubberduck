using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Refactoring.ParseTreeValue;
using Rubberduck.Refactorings;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete.UnreachableCaseEvaluation
{
    internal interface IParseTreeValueVisitor
    {
        IParseTreeVisitorResults VisitChildren(QualifiedModuleName module, IRuleNode node, DeclarationFinder finder);
    }

    internal class EnumMember
    {
        public EnumMember(VBAParser.EnumerationStmt_ConstantContext constContext, long initValue)
        {
            ConstantContext = constContext;
            Value = initValue;
            HasAssignment = constContext.children.Any(ch => ch.Equals(constContext.GetToken(VBAParser.EQ, 0)));
        }
        public VBAParser.EnumerationStmt_ConstantContext ConstantContext { get; }
        public long Value { set; get; }
        public bool HasAssignment { get; }
    }

    internal class ParseTreeValueVisitor : IParseTreeValueVisitor
    {
        private readonly IParseTreeValueFactory _valueFactory;
        private readonly Func<Declaration, (bool, string, string)> _valueDeclarationEvaluator;

        public ParseTreeValueVisitor(
            IParseTreeValueFactory valueFactory,
            Func<Declaration, (bool, string, string)> valueDeclarationEvaluator = null)
        {
            _valueFactory = valueFactory;
            _valueDeclarationEvaluator = valueDeclarationEvaluator ?? GetValuedDeclaration;
        }

        public IParseTreeVisitorResults VisitChildren(QualifiedModuleName module, IRuleNode ruleNode, DeclarationFinder finder)
        {
            var newResults = new ParseTreeVisitorResults();
            return VisitChildren(module, ruleNode, newResults, finder);
        }

        //The known results get passed along instead of aggregating from the bottom since other contexts can get already visited when resolving the value of other contexts.
        //Passing the results along avoids performing the resolution multiple times.
        private IMutableParseTreeVisitorResults VisitChildren(QualifiedModuleName module, IRuleNode node, IMutableParseTreeVisitorResults knownResults, DeclarationFinder finder)
        {
            if (!(node is ParserRuleContext context))
            {
                return knownResults;
            }

            var valueResults = knownResults;
            foreach (var child in context.children)
            {
                valueResults = Visit(module, child, valueResults, finder);
            }

            return valueResults;
        }

        private IMutableParseTreeVisitorResults Visit(QualifiedModuleName module, IParseTree tree, IMutableParseTreeVisitorResults knownResults, DeclarationFinder finder)
        {
            var valueResults = knownResults;
            if (tree is ParserRuleContext context && !(context is VBAParser.WhiteSpaceContext))
            {
                valueResults =  Visit(module, context, valueResults, finder);
            }

            return valueResults;
        }

        private IMutableParseTreeVisitorResults Visit(QualifiedModuleName module, ParserRuleContext parserRuleContext, IMutableParseTreeVisitorResults knownResults, DeclarationFinder finder)
        {
            switch (parserRuleContext)
            {
                case VBAParser.LExprContext lExpr:
                    return Visit(module, lExpr, knownResults, finder);
                case VBAParser.LiteralExprContext litExpr:
                    return Visit(litExpr, knownResults);
                case VBAParser.CaseClauseContext caseClause:
                    return VisitCaseClause(module, caseClause, knownResults, finder);
                case VBAParser.RangeClauseContext rangeClause:
                    return VisitRangeClause(module, rangeClause, knownResults, finder);
                case VBAParser.LogicalNotOpContext _:
                case VBAParser.UnaryMinusOpContext _:
                    return VisitUnaryOpEvaluationContext(module, parserRuleContext, knownResults, finder);
                default:
                    if (IsUnaryResultContext(parserRuleContext))
                    {
                        return VisitUnaryResultContext(module, parserRuleContext, knownResults, finder);
                    }
                    if (IsBinaryOpEvaluationContext(parserRuleContext))
                    {
                        return VisitBinaryOpEvaluationContext(module, parserRuleContext, knownResults, finder);
                    }

                    return knownResults;
            }
        }

        private IMutableParseTreeVisitorResults VisitRangeClause(
            QualifiedModuleName module, 
            VBAParser.RangeClauseContext rangeClause, 
            IMutableParseTreeVisitorResults knownResults, 
            DeclarationFinder finder)
        {
            var rangeClauseResults = VisitChildren(module, rangeClause, knownResults, finder);
            rangeClauseResults.AddIfNotPresent(rangeClause, _valueFactory.Create(rangeClause.GetText()));
            return rangeClauseResults;
        }

        private IMutableParseTreeVisitorResults VisitCaseClause(
            QualifiedModuleName module, 
            VBAParser.CaseClauseContext caseClause, 
            IMutableParseTreeVisitorResults knownResults,
            DeclarationFinder finder)
        {
            var caseClauseResults = VisitChildren(module, caseClause, knownResults, finder);
            caseClauseResults.AddIfNotPresent(caseClause, _valueFactory.Create(caseClause.GetText()));
            return caseClauseResults;
        }

        private IMutableParseTreeVisitorResults Visit(
            QualifiedModuleName module, 
            VBAParser.LExprContext context, 
            IMutableParseTreeVisitorResults knownResults, 
            DeclarationFinder finder)
        {
            if (knownResults.Contains(context))
            {
                return knownResults;
            }

            var valueResults = knownResults;

            IParseTreeValue newResult = null;
            if (TryGetLExprValue(module, context, ref valueResults, finder, out string lExprValue, out string declaredType))
            {
                newResult = _valueFactory.CreateDeclaredType(lExprValue, declaredType);
            }
            else
            {
                var simpleName = context.GetDescendent<VBAParser.SimpleNameExprContext>();
                if (TryGetIdentifierReferenceForContext(module, simpleName, finder, out var reference))
                {
                    var declarationTypeName = GetBaseTypeForDeclaration(reference.Declaration);
                    newResult = _valueFactory.CreateDeclaredType(context.GetText(), declarationTypeName);
                }
            }

            if (newResult != null)
            {
                valueResults.AddIfNotPresent(context, newResult);
            }

            return valueResults;
        }

        private IMutableParseTreeVisitorResults Visit(VBAParser.LiteralExprContext context, IMutableParseTreeVisitorResults knownResults)
        {
            if (knownResults.Contains(context))
            {
                return knownResults;
            }

            var valueResults = knownResults;
            var nResult = _valueFactory.Create(context.GetText());
            valueResults.AddIfNotPresent(context, nResult);

            return valueResults;
        }

        private IMutableParseTreeVisitorResults VisitBinaryOpEvaluationContext(
            QualifiedModuleName module, 
            ParserRuleContext context, 
            IMutableParseTreeVisitorResults knownResults, 
            DeclarationFinder finder)
        {
            var valueResults = VisitChildren(module, context, knownResults, finder);

            var (lhs, rhs, operatorSymbol) = RetrieveOpEvaluationElements(context, valueResults);
            if (lhs is null || rhs is null)
            {
                return valueResults;
            }
            if (lhs.IsOverflowExpression)
            {
                valueResults.AddIfNotPresent(context, lhs);
                return valueResults;
            }

            if (rhs.IsOverflowExpression)
            {
                valueResults.AddIfNotPresent(context, rhs);
                return valueResults;
            }

            var calculator = new ParseTreeExpressionEvaluator(_valueFactory, context.IsOptionCompareBinary());
            var result = calculator.Evaluate(lhs, rhs, operatorSymbol);
            valueResults.AddIfNotPresent(context, result);

            return valueResults;
        }

        private IMutableParseTreeVisitorResults VisitUnaryOpEvaluationContext(
            QualifiedModuleName module, 
            ParserRuleContext context, 
            IMutableParseTreeVisitorResults knownResults, 
            DeclarationFinder finder)
        {
            var valueResults = VisitChildren(module, context, knownResults, finder);

            var (lhs, rhs, operatorSymbol) = RetrieveOpEvaluationElements(context, valueResults);
            if (lhs is null || rhs != null)
            {
                return valueResults;
            }

            var calculator = new ParseTreeExpressionEvaluator(_valueFactory, context.IsOptionCompareBinary());
            var result = calculator.Evaluate(lhs, operatorSymbol);
            valueResults.AddIfNotPresent(context, result);

            return valueResults;
        }

        private static (IParseTreeValue LHS, IParseTreeValue RHS, string Symbol) RetrieveOpEvaluationElements(ParserRuleContext context, IMutableParseTreeVisitorResults knownResults)
        {
            (IParseTreeValue LHS, IParseTreeValue RHS, string Symbol) operandElements = (null, null, string.Empty);
            foreach (var child in NonWhitespaceChildren(context))
            {
                if (child is ParserRuleContext childContext)
                {
                    if (operandElements.LHS is null)
                    {
                        operandElements.LHS = knownResults.GetValue(childContext);
                    }
                    else if (operandElements.RHS is null)
                    {
                        operandElements.RHS = knownResults.GetValue(childContext);
                    }
                }
                else
                {
                    operandElements.Symbol = child.GetText();
                }
            }

            return operandElements;
        }

        private IMutableParseTreeVisitorResults VisitUnaryResultContext(
            QualifiedModuleName module, 
            ParserRuleContext parserRuleContext, 
            IMutableParseTreeVisitorResults knownResults, 
            DeclarationFinder finder)
        {
            var valueResults = VisitChildren(module, parserRuleContext, knownResults, finder);

            var firstChildWithValue = ParserRuleContextChildren(parserRuleContext)
                .FirstOrDefault(childContext => valueResults.Contains(childContext));

            if (firstChildWithValue != null)
            {
                valueResults.AddIfNotPresent(parserRuleContext, valueResults.GetValue(firstChildWithValue));
            }

            return valueResults;
        }

        private IMutableParseTreeVisitorResults VisitChildren(
            QualifiedModuleName module, 
            ParserRuleContext context, 
            IMutableParseTreeVisitorResults knownResults, 
            DeclarationFinder finder)
        {
            if (knownResults.Contains(context))
            {
                return knownResults;
            }

            var valueResults = knownResults;
            foreach (var childContext in ParserRuleContextChildren(context))
            {
                valueResults = Visit(module, childContext, valueResults, finder);
            }

            return valueResults;
        }

        private static IEnumerable<ParserRuleContext> ParserRuleContextChildren(ParserRuleContext ptParent)
            => NonWhitespaceChildren(ptParent).Where(ch => ch is ParserRuleContext).Cast<ParserRuleContext>();

        private static IEnumerable<IParseTree> NonWhitespaceChildren(ParserRuleContext ptParent)
            => ptParent.children.Where(ch => !(ch is VBAParser.WhiteSpaceContext));

        private bool TryGetLExprValue(QualifiedModuleName module, VBAParser.LExprContext lExprContext, ref IMutableParseTreeVisitorResults knownResults, DeclarationFinder finder, out string expressionValue, out string declaredTypeName)
        {
            expressionValue = string.Empty;
            declaredTypeName = string.Empty;
            if (lExprContext.TryGetChildContext(out VBAParser.MemberAccessExprContext memberAccess))
            {
                var member = memberAccess.GetChild<VBAParser.UnrestrictedIdentifierContext>();
                var (typeName, valueText, resultValues) = GetContextValue(module, member, knownResults, finder);
                knownResults = resultValues;
                declaredTypeName = typeName;
                expressionValue = valueText;
                return true;
            }

            if (lExprContext.TryGetChildContext(out VBAParser.SimpleNameExprContext smplName))
            {
                var (typeName, valueText, resultValues) = GetContextValue(module, smplName, knownResults, finder);
                knownResults = resultValues;
                declaredTypeName = typeName;
                expressionValue = valueText;
                return true;
            }

            if (lExprContext.TryGetChildContext(out VBAParser.IndexExprContext idxExpr)
                && ParseTreeValue.TryGetNonPrintingControlCharCompareToken(idxExpr.GetText(), out string comparableToken))
            {
                declaredTypeName = Tokens.String;
                expressionValue = comparableToken;
                return true;
            }

            return false;
        }

        private (bool IsType, string ExpressionValue, string TypeName) GetValuedDeclaration(Declaration declaration)
        {
            if (!(declaration is ValuedDeclaration valuedDeclaration))
            {
                return (false, null, null);
            }

            var typeName = GetBaseTypeForDeclaration(declaration);
            return (true, valuedDeclaration.Expression, typeName);
        }

        private (string declarationTypeName, string expressionValue, IMutableParseTreeVisitorResults resultValues) GetContextValue(
            QualifiedModuleName module, 
            ParserRuleContext context, 
            IMutableParseTreeVisitorResults knownResults, 
            DeclarationFinder finder)
        {
            if (!TryGetIdentifierReferenceForContext(module, context, finder, out var rangeClauseIdentifierReference))
            {
                return (string.Empty, context.GetText(), knownResults);
            }
            
            var declaration = rangeClauseIdentifierReference.Declaration;
            var expressionValue = rangeClauseIdentifierReference.IdentifierName;
            var declaredTypeName = GetBaseTypeForDeclaration(declaration);

            var (isValuedDeclaration, valuedExpressionValue, typeName) = _valueDeclarationEvaluator(declaration);
            if (isValuedDeclaration)
            {
                if (ParseTreeValue.TryGetNonPrintingControlCharCompareToken(valuedExpressionValue, out string resolvedValue))
                {
                    return (Tokens.String, resolvedValue, knownResults);
                }

                if (long.TryParse(valuedExpressionValue, out _))
                {
                    return (typeName, valuedExpressionValue, knownResults);
                }

                expressionValue = valuedExpressionValue;
                declaredTypeName = typeName;
            }

            if (declaration.DeclarationType.HasFlag(DeclarationType.Constant))
            {
                var (constantTokenExpressionValue, resultValues) = GetConstantContextValueToken(module, declaration.Context, knownResults, finder);
                return (declaredTypeName, constantTokenExpressionValue, resultValues);
            }

            if (declaration.DeclarationType.HasFlag(DeclarationType.EnumerationMember))
            {
                var (constantExpressionValue, resultValues) = GetConstantContextValueToken(module, declaration.Context, knownResults, finder);
                if (!constantExpressionValue.Equals(string.Empty))
                {
                    return (Tokens.Long, constantExpressionValue, resultValues);
                }

                if (declaration.Context.Parent is VBAParser.EnumerationStmtContext enumStmt)
                {
                    var (enumMembers, valueResults) = EnumMembers(module, enumStmt, resultValues, finder);
                    var enumValue = enumMembers.SingleOrDefault(enumMember => enumMember.ConstantContext == declaration.Context);
                    var enumExpressionValue = enumValue?.Value.ToString() ?? string.Empty;
                    return (Tokens.Long, enumExpressionValue, valueResults);
                } 

                return (Tokens.Long, string.Empty, resultValues);
            }

            return (declaredTypeName, expressionValue, knownResults);
        }

        private bool TryGetIdentifierReferenceForContext(
            QualifiedModuleName module, 
            ParserRuleContext context, 
            DeclarationFinder finder, 
            out IdentifierReference referenceForContext)
        {
            var (success, reference) = GetIdentifierReferenceForContext(module, context, finder);
            referenceForContext = reference;
            return success;
        }

        public static (bool success, IdentifierReference idRef) GetIdentifierReferenceForContext(
            QualifiedModuleName module, 
            ParserRuleContext context, 
            DeclarationFinder finder)
        {
            if (context == null)
            {
                return (false, null);
            }

            var qualifiedSelection = new QualifiedSelection(module, context.GetSelection());

            var identifierReferences =
                finder
                    .IdentifierReferences(qualifiedSelection)
                    .Where(reference => reference.Context == context)
                    .ToList();

            return identifierReferences.Count == 1
                ? (true, identifierReferences.First())
                : (false, null);
        }

        private (string valueText, IMutableParseTreeVisitorResults valueResults) GetConstantContextValueToken(
            QualifiedModuleName module, 
            ParserRuleContext context, 
            IMutableParseTreeVisitorResults knownResults,
            DeclarationFinder finder)
        {
            if (context is null)
            {
                return (string.Empty, knownResults);
            }

            var declarationContextChildren = context.children.ToList();
            var equalsSymbolIndex = declarationContextChildren.FindIndex(ch => ch.Equals(context.GetToken(VBAParser.EQ, 0)));

            var contextsOfInterest = new List<ParserRuleContext>();
            for (int idx = equalsSymbolIndex + 1; idx < declarationContextChildren.Count; idx++)
            {
                var childCtxt = declarationContextChildren[idx];
                if (!(childCtxt is VBAParser.WhiteSpaceContext))
                {
                    contextsOfInterest.Add((ParserRuleContext)childCtxt);
                }
            }

            foreach (var child in contextsOfInterest)
            {
                knownResults = Visit(module, child, knownResults, finder);
                if (knownResults.TryGetValue(child, out var value))
                {
                    return (value.Token, knownResults);
                }
            }
            return (string.Empty, knownResults);
        }

        private string GetBaseTypeForDeclaration(Declaration declaration)
        {
            var localDeclaration = declaration;
            var iterationGuard = 0;
            while (!(localDeclaration is null) 
                && !localDeclaration.AsTypeIsBaseType 
                && iterationGuard++ < 5)
            {
                localDeclaration = localDeclaration.AsTypeDeclaration;
            }
            return localDeclaration is null ? declaration.AsTypeName : localDeclaration.AsTypeName;
        }

        private static bool IsUnaryResultContext<T>(T context)
        {
            return context is VBAParser.SelectStartValueContext
                || context is VBAParser.SelectEndValueContext
                || context is VBAParser.ParenthesizedExprContext
                || context is VBAParser.SelectExpressionContext;
        }

        private static bool IsBinaryOpEvaluationContext<T>(T context)
        {
            if (context is VBAParser.ExpressionContext expressionContext)
            {

                return expressionContext.IsBinaryMathContext()
                    || expressionContext.IsBinaryLogicalContext()
                    || context is VBAParser.ConcatOpContext;
            }
            return false;
        }

        private (IReadOnlyList<EnumMember> enumMembers, IMutableParseTreeVisitorResults resultValues) EnumMembers(
            QualifiedModuleName enumModule, 
            VBAParser.EnumerationStmtContext enumerationStmtContext, 
            IMutableParseTreeVisitorResults knownResults,
            DeclarationFinder finder)
        {
            if (knownResults.TryGetEnumMembers(enumerationStmtContext, out var enumMembers))
            {
                return (enumMembers, knownResults);
            }

            var resultValues = LoadEnumMemberValues(enumModule, enumerationStmtContext, knownResults, finder);
            if (knownResults.TryGetEnumMembers(enumerationStmtContext, out var newEnumMembers))
            {
                return (newEnumMembers, resultValues);
            }

            return (new List<EnumMember>(), resultValues);
        }

        //The enum members incrementally to the parse tree visitor result are used within the call to Visit.
        private IMutableParseTreeVisitorResults LoadEnumMemberValues(
            QualifiedModuleName enumModule, 
            VBAParser.EnumerationStmtContext enumStmt, 
            IMutableParseTreeVisitorResults knownResults,
            DeclarationFinder finder)
        {
            var valueResults = knownResults;
            long enumAssignedValue = -1;
            var enumConstContexts = enumStmt.children
                .OfType<VBAParser.EnumerationStmt_ConstantContext>();
            foreach (var enumConstContext in enumConstContexts)
            {
                enumAssignedValue++;
                var enumMember = new EnumMember(enumConstContext, enumAssignedValue);
                if (enumMember.HasAssignment)
                {
                    valueResults = Visit(enumModule, enumMember.ConstantContext, valueResults, finder);

                    var (valueText, resultValues) = GetConstantContextValueToken(enumModule, enumMember.ConstantContext, valueResults, finder);
                    valueResults = resultValues;
                    if (!valueText.Equals(string.Empty))
                    {
                        enumMember.Value = long.Parse(valueText);
                        enumAssignedValue = enumMember.Value;
                    }
                }
                valueResults.AddEnumMember(enumStmt, enumMember);
            }

            return valueResults;
        }
    }
}
