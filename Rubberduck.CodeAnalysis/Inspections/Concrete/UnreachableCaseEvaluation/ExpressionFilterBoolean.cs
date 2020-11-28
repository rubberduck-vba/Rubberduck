using Rubberduck.Parsing.Grammar;
using Rubberduck.Refactoring.ParseTreeValue;
using Rubberduck.Refactorings;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete.UnreachableCaseEvaluation
{
    internal class ExpressionFilterBoolean : ExpressionFilter<bool>
    {
        public ExpressionFilterBoolean() : base(Tokens.Boolean, (a) => { return a.Equals(Tokens.True); }) { }

        public override IParseTreeValue SelectExpressionValue
        {
            set
            {
                _selectExpressionValue = value;
                AddSingleValue(!_selectExpressionValue.Token.Equals(Tokens.True));
            }
            get => _selectExpressionValue;
        }

        public override bool FiltersAllValues => FiltersTrueFalse;

        protected override bool AddIsClause(IsClauseExpression expression)
        {
            if (expression.LHS.ParsesToConstantValue)
            {
                return AddIsClause(expression.LHS.AsBoolean(), expression);
            }
            return AddToContainer(Variables[VariableClauseTypes.Is], expression.ToString());
        }

        protected override bool AddMinimum(bool value) { return false; }

        protected override bool AddMaximum(bool value) {return false; }

        protected override bool TryGetMaximum(out bool maximum) { maximum = default; return false; }

        protected override bool TryGetMinimum(out bool minimum) { minimum = default; return false; }

        protected override bool AddValueRange(RangeOfValues rov)
        {
            var addsStart = AddSingleValue(rov.Start);
            return AddSingleValue(rov.End) || addsStart;
        }

        private bool AddIsClause(bool isValue, IRangeClauseExpression expression)
        {
            /*
             * Indeterminant cases are added as comparable predicates
             * 
            *************************** Is Clause Boolean Truth Table  *********************
            * 
            *                           Select Expression Value
            *   Resolved Expression         True    False
            *   ****************************************************************************
            *   Is < True                   X       X       <= Inherently unreachable
            *   Is <= True                  True    False   
            *   Is > True                   False   True    
            *   Is >= True                  True    True    <= Matches for both True and False
            *   Is = True                   True    False
            *   Is <> True                  False   True
            *   Is > False                  X       X       <= Inherently unreachable
            *   Is >= False                 False   True
            *   Is < False                  True    False
            *   Is <= False                 True    True    <= Matches for both True and False
            *   Is = False                  False   True
            *   Is <> False                 True    False
            */

            var opSymbol = expression.OpSymbol;

            if (SelectExpressionValue is null)
            {
                if (opSymbol.Equals(RelationalOperators.NEQ))
                {
                    return AddSingleValue(!isValue);
                }

                if (opSymbol.Equals(RelationalOperators.EQ))
                {
                    return AddSingleValue(isValue);
                }

                if (opSymbol.Equals(RelationalOperators.GT))
                {
                    expression.IsInherentlyUnreachable = !isValue;
                    return isValue ? AddComparablePredicate(Tokens.Is, expression) : false;
                }

                if (opSymbol.Equals(RelationalOperators.GTE) || opSymbol.Equals(RelationalOperators.GTE2))
                {
                    return isValue ? AddTrueAndFalse()  //True for selectExpr value of True or False
                                        : AddComparablePredicate(Tokens.Is, expression);
                }

                if (opSymbol.Equals(RelationalOperators.LT))
                {
                    expression.IsInherentlyUnreachable = isValue;
                    return isValue ? false : AddComparablePredicate(Tokens.Is, expression);
                }

                if (opSymbol.Equals(RelationalOperators.LTE) || opSymbol.Equals(RelationalOperators.LTE2))
                {
                    return isValue ? AddComparablePredicate(Tokens.Is, expression)
                                        : AddTrueAndFalse(); //True for selectExpr value of True or False
                }
            }
            else //SelectExpressionContext resolves to True or False
            {
                var selectExpr = bool.Parse(SelectExpressionValue.Token);
                if (opSymbol.Equals(RelationalOperators.NEQ))
                {
                    return AddSingleValue(selectExpr != isValue);
                }

                if (opSymbol.Equals(RelationalOperators.EQ))
                {
                    return AddSingleValue(selectExpr == isValue);
                }

                if (opSymbol.Equals(RelationalOperators.GT))
                {
                    //if Is > True and the selectExpr is False => True
                    //If Is > True and the selectExpr is True => False
                    expression.IsInherentlyUnreachable = !isValue;
                    return isValue ? AddSingleValue(!selectExpr) : false;
                }

                if (opSymbol.Equals(RelationalOperators.GTE) || opSymbol.Equals(RelationalOperators.GTE2))
                {
                    //if Is >= True and the selectExpr is False => True
                    //If Is >= True and the selectExpr is True => True
                    //if Is >= False and the selectExpr is False => True
                    //If Is >= False and the selectExpr is True => False
                    return AddSingleValue(!(!isValue && selectExpr));
                }

                if (opSymbol.Equals(RelationalOperators.LT))
                {
                    expression.IsInherentlyUnreachable = isValue;
                    return isValue ? false : AddSingleValue(selectExpr);
                }

                if (opSymbol.Equals(RelationalOperators.LTE) || opSymbol.Equals(RelationalOperators.LTE2))
                {
                    //if Is <= True and the selectExpr is False => False
                    //If Is <= True and the selectExpr is True => True
                    //if Is <= False and the selectExpr is False => True
                    //If Is <= False and the selectExpr is True => True
                    return AddSingleValue(!(isValue && !selectExpr));
                }
            }

            return false;
        }

        private bool AddTrueAndFalse()
        {
            var addTrue = AddSingleValue(true);
            return AddSingleValue(false) || addTrue;
        }
    }
}
