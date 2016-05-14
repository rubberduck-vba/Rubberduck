﻿using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System;
using System.Diagnostics;

namespace Rubberduck.Parsing.Binding
{
    public sealed class BindingService
    {
        private readonly DeclarationFinder _declarationFinder;
        private readonly IBindingContext _defaultBindingContext;
        private readonly IBindingContext _typedBindingContext;
        private readonly IBindingContext _procedurePointerBindingContext;

        public BindingService(
            DeclarationFinder declarationFinder,
            IBindingContext defaultBindingContext,
            IBindingContext typedBindingContext,
            IBindingContext procedurePointerBindingContext)
        {
            _declarationFinder = declarationFinder;
            _defaultBindingContext = defaultBindingContext;
            _typedBindingContext = typedBindingContext;
            _procedurePointerBindingContext = procedurePointerBindingContext;
        }

        public Declaration ResolveEvent(Declaration module, string identifier)
        {
            return _declarationFinder.FindEvent(module, identifier);
        }

        public Declaration ResolveGoTo(Declaration procedure, string label)
        {
            return _declarationFinder.FindLabel(procedure, label);
        }

        public IBoundExpression ResolveDefault(Declaration module, Declaration parent, string expression, IBoundExpression withBlockVariable, ResolutionStatementContext statementContext)
        {
            var expr = Parse(expression.Trim());
            return _defaultBindingContext.Resolve(module, parent, expr, withBlockVariable, statementContext);
        }

        public IBoundExpression ResolveType(Declaration module, Declaration parent, string expression)
        {
            var expr = Parse(expression.Trim());
            return _typedBindingContext.Resolve(module, parent, expr, null, ResolutionStatementContext.Undefined);
        }

        public IBoundExpression ResolveProcedurePointer(Declaration module, Declaration parent, string expression)
        {
            var expr = Parse(expression.Trim());
            return _procedurePointerBindingContext.Resolve(module, parent, expr, null, ResolutionStatementContext.Undefined);
        }

        private ParserRuleContext Parse(string expression)
        {
            var stream = new AntlrInputStream(expression);
            var lexer = new VBALexer(stream);
            var tokens = new CommonTokenStream(lexer);
            var parser = new VBAExpressionParser(tokens);
            parser.AddErrorListener(new ExceptionErrorListener());
            var tree = parser.startRule();
            var prettyTree = tree.ToStringTree(parser);
            return tree;
        }
    }
}
