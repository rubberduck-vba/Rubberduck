using Rubberduck.Parsing.Grammar;
using System;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public class ExpressionFilterBoolean : ExpressionFilter<bool>
    {
        public ExpressionFilterBoolean(StringToValueConversion<bool> converter) : base(converter, Tokens.Boolean) { }

        public override IParseTreeValue SelectExpressionValue
        {
            set
            {
                _selectExpressionValue = value;
                AddSingleValue(!_selectExpressionValue.ValueText.Equals(Tokens.True));
            }
            get => _selectExpressionValue;
        }

        public override bool FiltersAllValues => FiltersTrueFalse;

        protected override bool AddIsClause(IsClauseExpression expression)
        {
            if (expression.LHSValue.ParsesToConstantValue)
            {
                if (Converter(expression.LHS, out bool bVal))
                {
                    return AddIsClause(bVal, expression);
                }
                throw new ArgumentException();
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
                if (opSymbol.Equals(LogicSymbols.NEQ))
                {
                    return AddSingleValue(!isValue);
                }
                else if (opSymbol.Equals(LogicSymbols.EQ))
                {
                    return AddSingleValue(isValue);
                }
                else if (opSymbol.Equals(LogicSymbols.GT))
                {
                    return isValue ? AddComparablePredicate(Tokens.Is, expression) : false;
                }
                else if (opSymbol.Equals(LogicSymbols.GTE))
                {
                    return isValue ? AddTrueAndFalse()  //True for selectExpr value of True or False
                                        : AddComparablePredicate(Tokens.Is, expression);
                }
                else if (opSymbol.Equals(LogicSymbols.LT))
                {
                    return isValue ? false : AddComparablePredicate(Tokens.Is, expression);
                }
                else if (opSymbol.Equals(LogicSymbols.LTE))
                {
                    return isValue ? AddComparablePredicate(Tokens.Is, expression)
                                        : AddTrueAndFalse(); //True for selectExpr value of True or False
                }
            }
            else //SelectExpressionContext resolves to True or False
            {
                var selectExpr = bool.Parse(SelectExpressionValue.ValueText);
                if (opSymbol.Equals(LogicSymbols.NEQ))
                {
                    return AddSingleValue(selectExpr != isValue);
                }
                else if (opSymbol.Equals(LogicSymbols.EQ))
                {
                    return AddSingleValue(selectExpr == isValue);
                }
                else if (opSymbol.Equals(LogicSymbols.GT))
                {
                    //if Is > True and the selectExpr is False => True
                    //If Is > True and the selectExpr is True => False
                    return isValue ? AddSingleValue(!selectExpr) : false;
                }
                else if (opSymbol.Equals(LogicSymbols.GTE))
                {
                    //if Is >= True and the selectExpr is False => True
                    //If Is >= True and the selectExpr is True => True
                    //if Is >= False and the selectExpr is False => True
                    //If Is >= False and the selectExpr is True => False
                    return AddSingleValue(!(!isValue && selectExpr));
                }
                else if (opSymbol.Equals(LogicSymbols.LT))
                {
                    return isValue ? false : AddSingleValue(selectExpr);
                }
                else if (opSymbol.Equals(LogicSymbols.LTE))
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
