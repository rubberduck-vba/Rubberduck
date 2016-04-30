using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class DefaultBindingContext : IBindingContext
    {
        private readonly DeclarationFinder _declarationFinder;

        public DefaultBindingContext(DeclarationFinder declarationFinder)
        {
            _declarationFinder = declarationFinder;
        }

        public IBoundExpression Resolve(Declaration module, Declaration parent, ParserRuleContext expression, ParserRuleContext innerMostWithVariableExpression)
        {
            dynamic dynamicExpression = expression;
            IExpressionBinding bindingTree = Visit(module, parent, dynamicExpression, innerMostWithVariableExpression);
            if (bindingTree != null)
            {
                return bindingTree.Resolve();
            }
            return null;
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.LExprContext expression, ParserRuleContext innerMostWithVariableExpression)
        {
            dynamic lexpr = expression.lExpression();
            return Visit(module, parent, lexpr, innerMostWithVariableExpression);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.NewExprContext expression, ParserRuleContext innerMostWithVariableExpression)
        {
            return Visit(module, parent, expression.newExpression(), innerMostWithVariableExpression);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.NewExpressionContext expression, ParserRuleContext innerMostWithVariableExpression)
        {
            var typeExpressionBinding = Visit(module, parent, expression.typeExpression(), innerMostWithVariableExpression);
            if (typeExpressionBinding == null)
            {
                return null;
            }
            return new NewTypeBinding(_declarationFinder, module, parent, expression, typeExpressionBinding);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.TypeExpressionContext expression, ParserRuleContext innerMostWithVariableExpression)
        {
            if (expression.builtInType() != null)
            {
                return null;
            }
            return Visit(module, parent, expression.definedTypeExpression(), innerMostWithVariableExpression);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.DefinedTypeExpressionContext expression, ParserRuleContext innerMostWithVariableExpression)
        {
            if (expression.simpleNameExpression() != null)
            {
                return VisitTypeContext(module, parent, expression.simpleNameExpression());
            }
            return VisitTypeContext(module, parent, expression.memberAccessExpression(), innerMostWithVariableExpression);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.SimpleNameExprContext expression, ParserRuleContext innerMostWithVariableExpression)
        {
            return Visit(module, parent, expression.simpleNameExpression(), innerMostWithVariableExpression);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.SimpleNameExpressionContext expression, ParserRuleContext innerMostWithVariableExpression)
        {
            return new SimpleNameDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.MemberAccessExprContext expression, ParserRuleContext innerMostWithVariableExpression)
        {
            dynamic lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression, innerMostWithVariableExpression);
            return new MemberAccessDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, lExpressionBinding);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.MemberAccessExpressionContext expression, ParserRuleContext innerMostWithVariableExpression)
        {
            dynamic lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression, innerMostWithVariableExpression);
            return new MemberAccessDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, lExpressionBinding);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.IndexExprContext expression, ParserRuleContext innerMostWithVariableExpression)
        {
            dynamic lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression, innerMostWithVariableExpression);
            return new IndexDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, lExpressionBinding);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.IndexExpressionContext expression, ParserRuleContext innerMostWithVariableExpression)
        {
            dynamic lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression, innerMostWithVariableExpression);
            return new IndexDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, lExpressionBinding);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.DictionaryAccessExprContext expression, ParserRuleContext innerMostWithVariableExpression)
        {
            dynamic lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression, innerMostWithVariableExpression);
            return VisitDictionaryAccessExpression(module, parent, expression, lExpressionBinding);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.DictionaryAccessExpressionContext expression, ParserRuleContext innerMostWithVariableExpression)
        {
            dynamic lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression, innerMostWithVariableExpression);
            return VisitDictionaryAccessExpression(module, parent, expression, lExpressionBinding);
        }

        private IExpressionBinding VisitDictionaryAccessExpression(Declaration module, Declaration parent, ParserRuleContext expression, IExpressionBinding lExpressionBinding)
        {
            /*
                A dictionary access expression is syntactically translated into an index expression with the same 
                expression for <l-expression> and an argument list with a single positional argument with a 
                declared type of String and a value equal to the name value of <unrestricted-name>. 
             */
            var fakeArgList = new ArgumentList();
            fakeArgList.AddArgument(ArgumentListArgumentType.Positional);
            return new IndexDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, lExpressionBinding, fakeArgList);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.WithExprContext expression, ParserRuleContext innerMostWithVariableExpression)
        {
            return Visit(module, parent, expression.withExpression(), innerMostWithVariableExpression);
        }

        private IExpressionBinding Visit(Declaration module, Declaration parent, VBAExpressionParser.WithExpressionContext expression, ParserRuleContext innerMostWithVariableExpression)
        {
            /*
                A <with-member-access-expression> or <with-dictionary-access-expression> is 
                statically resolved as a normal member access or dictionary access expression, respectively, as if 
                the innermost enclosing With block variable was specified for <l-expression>. If there is no 
                enclosing With block, the <with-expression> is invalid.
             */
            if (innerMostWithVariableExpression == null)
            {
                return null;
            }
            dynamic lExpression = innerMostWithVariableExpression;
            var lExpressionBinding = Visit(module, parent, lExpression, null);
            if (expression.withMemberAccessExpression() != null)
            {
                return new MemberAccessDefaultBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, innerMostWithVariableExpression, lExpressionBinding, expression.withMemberAccessExpression().unrestrictedName().GetText());
            }
            else
            {
                return VisitDictionaryAccessExpression(module, parent, expression, lExpressionBinding);
            }
        }

        private IExpressionBinding VisitTypeContext(Declaration module, Declaration parent, VBAExpressionParser.SimpleNameExprContext expression)
        {
            return VisitTypeContext(module, parent, expression.simpleNameExpression());
        }

        private IExpressionBinding VisitTypeContext(Declaration module, Declaration parent, VBAExpressionParser.SimpleNameExpressionContext expression)
        {
            return new SimpleNameTypeBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression);
        }

        private IExpressionBinding VisitTypeContext(Declaration module, Declaration parent, VBAExpressionParser.MemberAccessExprContext expression, ParserRuleContext innerMostWithVariableExpression)
        {
            dynamic lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression, innerMostWithVariableExpression);
            return new MemberAccessTypeBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, lExpressionBinding);
        }

        private IExpressionBinding VisitTypeContext(Declaration module, Declaration parent, VBAExpressionParser.MemberAccessExpressionContext expression, ParserRuleContext innerMostWithVariableExpression)
        {
            dynamic lExpression = expression.lExpression();
            var lExpressionBinding = Visit(module, parent, lExpression, innerMostWithVariableExpression);
            return new MemberAccessTypeBinding(_declarationFinder, Declaration.GetProjectParent(parent), module, parent, expression, lExpressionBinding);
        }
    }
}
