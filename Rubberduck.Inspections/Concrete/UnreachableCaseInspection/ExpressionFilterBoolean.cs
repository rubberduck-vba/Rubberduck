using Rubberduck.Parsing.Grammar;
using System;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public class ExpressionFilterBoolean : ExpressionFilter<bool>
    {
        public ExpressionFilterBoolean(StringToValueConversion<bool> converter) : base(converter, Tokens.Boolean) { }

        protected override bool AddIsClause(IsClauseExpression expression)
        {
            if (expression.LHSValue.ParsesToConstantValue)
            {
                //if(expression.LHSValue.TryConvertValue(out bool bVal))
                if (Converter(expression.LHS, Tokens.Boolean, out bool bVal))
                {
                    return AddIsClause(bVal, expression);
                }
                throw new ArgumentException();
            }
            return AddToContainer(Variables[VariableClauseTypes.Is], expression.ToString());
        }

        private bool AddIsClause(bool val, IRangeClauseExpression expression)
        {
            /*
             * Indeterminant cases are added as comparable predicates
             * 
            *************************** Is Clause Boolean Truth Table  *********************
            * 
            *                          Select Case Value
            *   Resolved Expression     True    False
            *   ****************************************************************************
            *   Is < True               False   False   <= Always False
            *   Is <= True              True    False   
            *   Is > True               False   True    
            *   Is >= True              True    True    <= Always True
            *   Is = True               True    False
            *   Is <> True              False   True
            *   Is > False              False   False   <= Always False
            *   Is >= False             False   True
            *   Is < False              True    False
            *   Is <= False             True    True    <= Always True
            *   Is = False              False   True
            *   Is <> False             True    False
            */

            bool bVal = val.CompareTo(true) == 0;
            var opSymbol = expression.OpSymbol;

            if (opSymbol.Equals(LogicSymbols.NEQ)
                || opSymbol.Equals(LogicSymbols.EQ)
                || (opSymbol.Equals(LogicSymbols.GT) && bVal)
                || (opSymbol.Equals(LogicSymbols.LT) && !bVal)
                || (opSymbol.Equals(LogicSymbols.GTE) && !bVal)
                || (opSymbol.Equals(LogicSymbols.LTE) && bVal)
                )
            {
                var inputPredicate = new PredicateValueExpression<string>(Tokens.Is, val.ToString(), opSymbol);
                AddComparablePredicate(Tokens.Is, val.ToString(), opSymbol);
            }
            else if (opSymbol.Equals(LogicSymbols.GT) || opSymbol.Equals(LogicSymbols.GTE))
            {
                return AddSingleValue(bVal);
            }
            else if (opSymbol.Equals(LogicSymbols.LT) || opSymbol.Equals(LogicSymbols.LTE))
            {
                return AddSingleValue(!bVal);
            }
            return false;
        }

        protected override bool AddMinimum(bool value) { return false; }

        protected override bool AddMaximum(bool value) {return false; }

        protected override bool TryGetMaximum(out bool maximum) { maximum = default; return false; }

        protected override bool TryGetMinimum(out bool minimum) { minimum = default; return false; }

        //protected override bool AddValueRange(RangeValues<bool> range)
        protected override bool AddValueRange((bool Start, bool End) range)
        {
            var addsStart = AddSingleValue(range.Start);
            var addsEnd = AddSingleValue(range.End);
            return addsStart || addsEnd;
        }

        public override bool FiltersAllValues => FiltersTrueFalse;
    }
}
