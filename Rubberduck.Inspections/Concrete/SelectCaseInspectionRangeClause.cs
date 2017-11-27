using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.Inspections.Concrete
{
    internal static class CompareSymbols
    {
        public static readonly string EQ = "=";
        public static readonly string NEQ = "<>";
        public static readonly string LT = "<";
        public static readonly string LTE = "<=";
        public static readonly string GT = ">";
        public static readonly string GTE = ">=";
    }

    internal static class CompareExtents
    {
        public static long LONGMIN = -2147486648;
        public static long LONGMAX = 2147486647;
        public static long INTEGERMIN = -32768;
        public static long INTEGERMAX = 32767;
        public static long BYTEMIN = 0;
        public static long BYTEMAX = 255;
        public static decimal CURRENCYMIN = -922337203685477.5808M;
        public static decimal CURRENCYMAX = 922337203685477.5807M;
        public static double SINGLEMIN = -3402823E38;
        public static double SINGLEMAX = 3402823E38;
    }

    //public class SelectCaseInspectionRangeClause
    //{
    //    private readonly VBAParser.RangeClauseContext _ctxt;
    //    private readonly RubberduckParserState _state;
    //    private readonly string _idReferenceName;
    //    private bool _usesIsClause;
    //    private bool _isValueRange;
    //    private bool _isParseable;
    //    private string _nativeTypeName;
    //    private string _targetTypeName;
    //    private string _compareSymbol;
    //    private SelectCaseInspectionValue _minValue;
    //    private SelectCaseInspectionValue _maxValue;

    //    internal SelectCaseInspectionRangeClause(RubberduckParserState state, string typeName, string idReferenceName, VBAParser.RangeClauseContext ctxt)
    //    {
    //        _ctxt = ctxt;
    //        _usesIsClause = HasChildToken(ctxt, Tokens.Is);
    //        _isValueRange = HasChildToken(ctxt, Tokens.To);
    //        _compareSymbol = _usesIsClause ? GetTheCompareOperator(ctxt) : CompareSymbols.EQ;
    //        _targetTypeName = typeName;
    //        _state = state;
    //        _idReferenceName = idReferenceName;
    //        _nativeTypeName = EvaluateRangeClauseTypeName(ctxt);

    //        if (_isValueRange)
    //        {
    //            var startValueAsString = GetText(ParserRuleContextHelper.GetChild<VBAParser.SelectStartValueContext>(_ctxt));
    //            var endValueAsString = GetText(ParserRuleContextHelper.GetChild<VBAParser.SelectEndValueContext>(_ctxt));

    //            var startValue = new SelectCaseInspectionValue(startValueAsString, TargetTypename);
    //            var endValue = new SelectCaseInspectionValue(endValueAsString, TargetTypename);

    //            _minValue = startValue <= endValue ? startValue : endValue;
    //            _maxValue = startValue <= endValue ? endValue : startValue;

    //            _isParseable = _minValue.IsParseable && _maxValue.IsParseable;
    //        }
    //        else
    //        {
    //            _minValue = new SelectCaseInspectionValue(GetRangeClauseText(_ctxt), TargetTypename);
    //            _maxValue = _minValue;
    //            _isParseable = _minValue.IsParseable;
    //        }

    //        CompareByTextOnly = !IsParseable && NativeTypeName.Equals(TargetTypename);
    //    }

    //    public string TargetTypename => _targetTypeName;

    //    public bool IsParseable => _isParseable;
    //    public bool CompareByTextOnly { set; get; }
    //    public bool MatchesSelectCaseType => NativeTypeName.Equals(TargetTypename);
    //    public string NativeTypeName => _nativeTypeName;
    //    public VBAParser.RangeClauseContext Context => _ctxt;

    //    public bool IsSingleVal => !IsRange;
    //    public bool IsRange => _isValueRange;
    //    public bool UsesIsClause => _usesIsClause;
    //    public SelectCaseInspectionValue SingleValue => _minValue;
    //    public SelectCaseInspectionValue MinValue => _minValue;
    //    public SelectCaseInspectionValue MaxValue => _maxValue;
    //    public string CompareSymbol => _compareSymbol;

    //    public bool HasOutOfBoundsValue { set; get; }
    //    public bool IsPreviouslyHandled { set; get; }
    //    public bool CausesUnreachableCaseElse { set; get; }

    //    private bool IsSelectCaseBoolean => TargetTypename.Equals(Tokens.Boolean);
    //    private bool IsStringLiteral(string text) => text.StartsWith("\"") && _ctxt.GetText().EndsWith("\"");

    //    private static Dictionary<string, string> _compareInversions = new Dictionary<string, string>()
    //    {
    //        { CompareSymbols.EQ, CompareSymbols.NEQ },
    //        { CompareSymbols.NEQ, CompareSymbols.EQ },
    //        { CompareSymbols.LT, CompareSymbols.GT },
    //        { CompareSymbols.LTE, CompareSymbols.GTE },
    //        { CompareSymbols.GT, CompareSymbols.LT },
    //        { CompareSymbols.GTE, CompareSymbols.LTE }
    //    };

    //    private static Dictionary<string, string> _compareInversionsExtended = new Dictionary<string, string>()
    //    {
    //        { CompareSymbols.EQ, CompareSymbols.NEQ },
    //        { CompareSymbols.NEQ, CompareSymbols.EQ },
    //        { CompareSymbols.LT, CompareSymbols.GTE },
    //        { CompareSymbols.LTE, CompareSymbols.GT },
    //        { CompareSymbols.GT, CompareSymbols.LTE },
    //        { CompareSymbols.GTE, CompareSymbols.LT }
    //    };

    //    private static bool IsComparisonOperator(string opCandidate) => _compareInversions.Keys.Contains(opCandidate);
    //    private static string GetInverse(string theOperator)
    //    {
    //        return IsComparisonOperator(theOperator) ? _compareInversions[theOperator] : theOperator;
    //    }

    //    private string EvaluateRangeClauseTypeName(VBAParser.RangeClauseContext rangeCtxt)
    //    {
    //        var textValue = rangeCtxt.GetText();
    //        if (IsStringLiteral(textValue))
    //        {
    //            return Tokens.String;
    //        }
    //        else if (textValue.EndsWith("#"))
    //        {
    //            var modified = textValue.Substring(0,textValue.Length - 1);
    //            if (long.TryParse(modified, out _ ))
    //            {
    //                return Tokens.Double;
    //            }
    //            return Tokens.String;
    //        }
    //        else if (textValue.Contains("."))
    //        {
    //            if(double.TryParse(textValue, out _))
    //            {
    //                return Tokens.Double;
    //            }

    //            if (decimal.TryParse(textValue, out _))
    //            {
    //                return Tokens.Currency;
    //            }
    //            return TargetTypename;
    //        }
    //        else if (textValue.Equals(Tokens.True) || textValue.Equals(Tokens.False))
    //        {
    //            return Tokens.Boolean;
    //        }
    //        else
    //        {
    //            return TargetTypename;
    //        }
    //    }

    //    private IdentifierReference GetTheRangeClauseReference(ParserRuleContext rangeClauseCtxt, string theName)
    //    {
    //        var allRefs = new List<IdentifierReference>();
    //        foreach (var dec in _state.DeclarationFinder.MatchName(theName))
    //        {
    //            allRefs.AddRange(dec.References);
    //        }

    //        if (!allRefs.Any())
    //        {
    //            return null;
    //        }

    //        if (allRefs.Count == 1)
    //        {
    //            return allRefs.First();
    //        }
    //        else
    //        {
    //            var simpleNameExpr = ParserRuleContextHelper.GetChild<VBAParser.SimpleNameExprContext>(rangeClauseCtxt);
    //            var rangeClauseReference = allRefs.Where(rf => ParserRuleContextHelper.HasParent(rf.Context, rangeClauseCtxt)
    //                                    && (ParserRuleContextHelper.HasParent(rf.Context, simpleNameExpr.Parent)));

    //            Debug.Assert(rangeClauseReference.Count() == 1);
    //            return rangeClauseReference.First();
    //        }
    //    }

    //    private string GetRangeClauseText(VBAParser.RangeClauseContext ctxt)
    //    {
    //        VBAParser.RelationalOpContext relationalOpCtxt;
    //        if (TryGetChildContext(ctxt, out relationalOpCtxt))
    //        {
    //            _usesIsClause = true;
    //            return GetTextForRelationalOpContext(relationalOpCtxt);
    //        }

    //        VBAParser.LExprContext lExprContext;
    //        if (TryGetChildContext(ctxt, out lExprContext))
    //        {
    //            string expressionValue;
    //            return TryGetTheExpressionValue(lExprContext, out expressionValue) ? expressionValue : string.Empty;
    //        }

    //        VBAParser.UnaryMinusOpContext negativeCtxt;
    //        if (TryGetChildContext(ctxt, out negativeCtxt))
    //        {
    //            return negativeCtxt.GetText();
    //        }

    //        VBAParser.LiteralExprContext theValCtxt;
    //        return TryGetChildContext(ctxt, out theValCtxt) ? GetText(theValCtxt) : string.Empty;
    //    }

    //    private string GetTextForRelationalOpContext(VBAParser.RelationalOpContext relationalOpCtxt)
    //    {
    //        var lExprCtxtIndices = new List<int>();
    //        var literalExprCtxtIndices = new List<int>();

    //        for (int idx = 0; idx < relationalOpCtxt.ChildCount; idx++)
    //        {
    //            var text = relationalOpCtxt.children[idx].GetText();
    //            if (relationalOpCtxt.children[idx] is VBAParser.LExprContext)
    //            {
    //                lExprCtxtIndices.Add(idx);
    //            }
    //            else if (relationalOpCtxt.children[idx] is VBAParser.UnaryMinusOpContext
    //                    || relationalOpCtxt.children[idx] is VBAParser.LiteralExprContext)
    //            {
    //                literalExprCtxtIndices.Add(idx);
    //            }
    //            else if (IsComparisonOperator(text))
    //            {
    //                _compareSymbol = text;
    //            }
    //        }

    //        if (lExprCtxtIndices.Count() == 2)  //e.g., x > someConstantExpression
    //        {
    //            var ctxtLHS = (VBAParser.LExprContext)relationalOpCtxt.children[lExprCtxtIndices.First()];
    //            var ctxtRHS = (VBAParser.LExprContext)relationalOpCtxt.children[lExprCtxtIndices.Last()];

    //            string result;
    //            if (GetText(ctxtLHS).Equals(_idReferenceName))
    //            {
    //                return TryGetTheExpressionValue(ctxtRHS, out result) ? result : string.Empty;
    //            }
    //            else if (GetText(ctxtRHS).Equals(_idReferenceName))
    //            {
    //                _compareSymbol = GetInverse(_compareSymbol);
    //                return TryGetTheExpressionValue(ctxtLHS, out result) ? result : string.Empty;
    //            }
    //        }
    //        else if (lExprCtxtIndices.Count == 1 && literalExprCtxtIndices.Count == 1) // e.g., z < 10
    //        {
    //            var lExpIndex = lExprCtxtIndices.First();
    //            var litExpIndex = literalExprCtxtIndices.First();
    //            var lExprCtxt = (VBAParser.LExprContext)relationalOpCtxt.children[lExpIndex];
    //            if (GetText(lExprCtxt).Equals(_idReferenceName))
    //            {
    //                _compareSymbol = lExpIndex > litExpIndex ?
    //                    GetInverse(_compareSymbol) : _compareSymbol;
    //                return GetText((ParserRuleContext)relationalOpCtxt.children[litExpIndex]);
    //            }
    //        }
    //        return string.Empty;
    //    }

    //    private bool TryGetTheExpressionValue(VBAParser.LExprContext ctxt, out string expressionValue)
    //    {
    //        expressionValue = string.Empty;
    //        var smplName = ParserRuleContextHelper.GetDescendent<VBAParser.SimpleNameExprContext>(ctxt);
    //        if (smplName != null)
    //        {
    //            var rangeClauseIdentifierReference = GetTheRangeClauseReference(smplName, smplName.GetText());
    //            if (rangeClauseIdentifierReference != null)
    //            {
    //                if (rangeClauseIdentifierReference.Declaration.DeclarationType.HasFlag(DeclarationType.Constant))
    //                {
    //                    var valuedDeclaration = (ConstantDeclaration)rangeClauseIdentifierReference.Declaration;
    //                    expressionValue = valuedDeclaration.Expression;
    //                    return true;
    //                }
    //            }
    //        }
    //        return false;
    //    }

    //    private string GetText(ParserRuleContext ctxt)
    //    {
    //        var text = ctxt.GetText();
    //        return text.Replace("\"", "");
    //    }

    //    private string GetTheCompareOperator(VBAParser.RangeClauseContext ctxt)
    //    {
    //        VBAParser.ComparisonOperatorContext opCtxt;
    //        _usesIsClause = TryGetChildContext(ctxt, out opCtxt);
    //        return _usesIsClause ? opCtxt.GetText() : CompareSymbols.EQ;
    //    }

    //    private static bool HasChildToken<T>(T ctxt, string token) where T : ParserRuleContext
    //    {
    //        var result = false;
    //        for (int idx = 0; idx < ctxt.ChildCount && !result; idx++)
    //        {
    //            if (ctxt.children[idx].GetText().Equals(token))
    //            {
    //                result = true;
    //            }
    //        }
    //        return result;
    //    }

    //    private static bool TryGetChildContext<T, U>(T ctxt, out U opCtxt) where T : ParserRuleContext where U : ParserRuleContext //VBAParser.ExpressionContext
    //    {
    //        opCtxt = null;
    //        opCtxt = ParserRuleContextHelper.GetChild<U>(ctxt);
    //        return opCtxt != null;
    //    }
    //}

    #region SelectCaseInspectionValue
    public class SelectCaseInspectionValue
    {
        private readonly string _targetTypeName;
        private readonly string _valueAsString;
        private readonly Func<SelectCaseInspectionValue, SelectCaseInspectionValue, bool> _operatorIsGT;
        private readonly Func<SelectCaseInspectionValue, SelectCaseInspectionValue, bool> _operatorIsLT;
        private readonly Func<SelectCaseInspectionValue, SelectCaseInspectionValue, bool> _operatorIsEQ;

        private long? _valueAsLong;
        private double? _valueAsDouble;
        private decimal? _valueAsDecimal;
        private bool? _valueAsBoolean;

        private long resultLong;
        private double resultDouble;
        private decimal resultDecimal;

        private static Dictionary<string, string> UpperBounds = new Dictionary<string, string>()
        {
            { Tokens.Integer, CompareExtents.INTEGERMAX.ToString()},
            { Tokens.Long, CompareExtents.LONGMAX.ToString()},
            { Tokens.Byte, CompareExtents.BYTEMAX.ToString()},
            { Tokens.Currency, CompareExtents.CURRENCYMAX.ToString()},
            { Tokens.Single, CompareExtents.SINGLEMAX.ToString()}
        };

        private static Dictionary<string, string> LowerBounds = new Dictionary<string, string>()
        {
            { Tokens.Integer, CompareExtents.INTEGERMIN.ToString()},
            { Tokens.Long, CompareExtents.LONGMIN.ToString()},
            { Tokens.Byte, CompareExtents.BYTEMIN.ToString()},
            { Tokens.Currency, CompareExtents.CURRENCYMIN.ToString()},
            { Tokens.Single, CompareExtents.SINGLEMIN.ToString()}
        };

        private static Dictionary<string, Func<SelectCaseInspectionValue, SelectCaseInspectionValue, bool>> OperatorsIsGT = new Dictionary<string, Func<SelectCaseInspectionValue, SelectCaseInspectionValue, bool>>()
        {
            { Tokens.Integer, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsLong().Value > compValue.AsLong().Value; } },
            { Tokens.Long, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsLong().Value > compValue.AsLong().Value; } },
            { Tokens.Byte, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsLong().Value > compValue.AsLong().Value; } },
            { Tokens.Double, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsDouble().Value > compValue.AsDouble().Value; } },
            { Tokens.Single, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsDouble().Value > compValue.AsDouble().Value; } },
            { Tokens.Currency, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsCurrency().Value > compValue.AsCurrency().Value; } },
            { Tokens.Boolean, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsBoolean().Value == compValue.AsBoolean().Value ? false : !thisValue.AsBoolean().Value; } },
            { Tokens.String, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsString().CompareTo(compValue.AsString()) > 0; } }
        };

        private static Dictionary<string, Func<SelectCaseInspectionValue, SelectCaseInspectionValue, bool>> OperatorsIsLT = new Dictionary<string, Func<SelectCaseInspectionValue, SelectCaseInspectionValue, bool>>()
        {
            { Tokens.Integer, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsLong().Value < compValue.AsLong().Value; } },
            { Tokens.Long, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsLong().Value < compValue.AsLong().Value; } },
            { Tokens.Byte, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsLong().Value < compValue.AsLong().Value; } },
            { Tokens.Double, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsDouble().Value < compValue.AsDouble().Value; } },
            { Tokens.Single, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsDouble().Value < compValue.AsDouble().Value; } },
            { Tokens.Currency, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsCurrency().Value < compValue.AsCurrency().Value; } },
            { Tokens.Boolean, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsBoolean().Value == compValue.AsBoolean().Value ? false : thisValue.AsBoolean().Value; } },
            { Tokens.String, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsString().CompareTo(compValue.AsString()) < 0; } }
        };

        private static Dictionary<string, Func<SelectCaseInspectionValue, SelectCaseInspectionValue, bool>> OperatorsIsEQ = new Dictionary<string, Func<SelectCaseInspectionValue, SelectCaseInspectionValue, bool>>()
        {
            { Tokens.Integer, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsLong().Value == compValue.AsLong().Value; } },
            { Tokens.Long, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsLong().Value == compValue.AsLong().Value; } },
            { Tokens.Byte, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsLong().Value == compValue.AsLong().Value; } },
            { Tokens.Double, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsDouble().Value == compValue.AsDouble().Value; } },
            { Tokens.Single, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsDouble().Value == compValue.AsDouble().Value; } },
            { Tokens.Currency, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsCurrency().Value == compValue.AsCurrency().Value; } },
            { Tokens.Boolean, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsBoolean().Value == compValue.AsBoolean().Value; } },
            { Tokens.String, delegate(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue){ return thisValue.AsString().CompareTo(compValue.AsString()) == 0; } }
        };

        private static Dictionary<string, Func<SelectCaseInspectionValue, bool>> IsParseableTests = new Dictionary<string, Func<SelectCaseInspectionValue, bool>>()
        {
            { Tokens.Integer, delegate(SelectCaseInspectionValue thisValue){ return thisValue.AsLong().HasValue; } },
            { Tokens.Long, delegate(SelectCaseInspectionValue thisValue){ return thisValue.AsLong().HasValue; } },
            { Tokens.Byte, delegate(SelectCaseInspectionValue thisValue){ return thisValue.AsLong().HasValue; } },
            { Tokens.Double, delegate(SelectCaseInspectionValue thisValue){ return thisValue.AsDouble().HasValue; } },
            { Tokens.Single, delegate(SelectCaseInspectionValue thisValue){ return thisValue.AsDouble().HasValue; } },
            { Tokens.Currency, delegate(SelectCaseInspectionValue thisValue){ return thisValue.AsCurrency().HasValue; } },
            { Tokens.Boolean, delegate(SelectCaseInspectionValue thisValue){ return thisValue.AsBoolean().HasValue; } },
            { Tokens.String, delegate(SelectCaseInspectionValue thisValue){ return true; } }
        };

        public SelectCaseInspectionValue(string valueToken, string targetTypeName)
        {
            _valueAsString = valueToken.EndsWith("#") ? valueToken.Replace("#", ".00") : valueToken;
            _targetTypeName = targetTypeName;

            Debug.Assert(OperatorsIsGT.ContainsKey(targetTypeName));
            Debug.Assert(OperatorsIsLT.ContainsKey(targetTypeName));
            Debug.Assert(OperatorsIsEQ.ContainsKey(targetTypeName));

            _operatorIsGT = OperatorsIsGT[targetTypeName];
            _operatorIsLT = OperatorsIsLT[targetTypeName];
            _operatorIsEQ = OperatorsIsEQ[targetTypeName];
        }

        public static SelectCaseInspectionValue CreateUpperBound(string typename)
        {
            if (UpperBounds.ContainsKey(typename))
            {
                return new SelectCaseInspectionValue(UpperBounds[typename], typename);
            }
            return null;
        }

        public static SelectCaseInspectionValue CreateLowerBound(string typename)
        {
            if (LowerBounds.ContainsKey(typename))
            {
                return new SelectCaseInspectionValue(LowerBounds[typename], typename);
            }
            return null;
        }

        public bool IsIntegerNumber => new string[] { Tokens.Long, Tokens.Integer, Tokens.Byte }.Contains(TargetTypeName);

        public string TargetTypeName => _targetTypeName;

        public bool IsParseable
            => IsParseableTests.ContainsKey(TargetTypeName) ? IsParseableTests[TargetTypeName](this) : false;

        public bool IsWithin(SelectCaseInspectionValue start, SelectCaseInspectionValue end ) 
            => start > end ? this >= end && this <= start : this >= start && this <= end;


        public static bool operator >(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue)
        {
            return thisValue._operatorIsGT(thisValue, compValue);
        }

        public static bool operator <(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue)
        {
            return thisValue._operatorIsLT(thisValue, compValue);
        }

        public static bool operator ==(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue)
        {
            if(ReferenceEquals(null, thisValue))
            {
                return ReferenceEquals(null, compValue);
            }
            else
            {
                return ReferenceEquals(null, compValue) ? false : thisValue._operatorIsEQ(thisValue, compValue);
            }
        }

        public static bool operator !=(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue)
        {
            if (ReferenceEquals(null, thisValue))
            {
                return !ReferenceEquals(null, compValue);
            }
            else
            {
                return ReferenceEquals(null, compValue) ? true : !(thisValue == compValue);
            }
        }

        public static bool operator >=(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue)
        {
            return thisValue == compValue || thisValue > compValue;
        }

        public static bool operator <=(SelectCaseInspectionValue thisValue, SelectCaseInspectionValue compValue)
        {
            return thisValue == compValue || thisValue < compValue;
        }

        public override bool Equals(Object obj)
        {
            if (ReferenceEquals(null, obj) || !(obj is SelectCaseInspectionValue))
            {
                return false;
            }
            var asValue = (SelectCaseInspectionValue)obj;
            return asValue.TargetTypeName == TargetTypeName ? asValue == this : false;
        }

        public override int GetHashCode()
        {
            return _valueAsString.GetHashCode();
        }

        public string AsString()
        {
            return _valueAsString;
        }

        public long? AsLong()
        {
            if (!_valueAsLong.HasValue)
            {
                if (long.TryParse(_valueAsString, out resultLong))
                {
                    _valueAsLong = resultLong;
                }
                else if (decimal.TryParse(_valueAsString, out resultDecimal))
                {
                    _valueAsLong = SafeConvertToLong(resultDecimal);
                }
                else if (double.TryParse(_valueAsString, out resultDouble))
                {
                    _valueAsLong = SafeConvertToLong(resultDouble);
                }
                else if (_valueAsString.Equals(Tokens.True))
                {
                    _valueAsLong = -1;
                }
                else if (_valueAsString.Equals(Tokens.False))
                {
                    _valueAsLong = 0;
                }
            }
            return _valueAsLong;
        }

        public decimal? AsCurrency()
        {
            if (!_valueAsDecimal.HasValue)
            {
                if (decimal.TryParse(_valueAsString, out resultDecimal))
                {
                    _valueAsDecimal = resultDecimal;
                }
                else if (double.TryParse(_valueAsString, out resultDouble))
                {
                    _valueAsDecimal = SafeConvertToDecimal(resultDouble);
                }
                else if (long.TryParse(_valueAsString, out resultLong))
                {
                    _valueAsDecimal = SafeConvertToDecimal(resultLong);
                }
                else if (_valueAsString.Equals(Tokens.True))
                {
                    _valueAsDecimal = -1.0M;
                }
                else if (_valueAsString.Equals(Tokens.False))
                {
                    _valueAsDecimal = 0.0M;
                }
            }
            return _valueAsDecimal;
        }

        public double? AsDouble()
        {
            if (!_valueAsDouble.HasValue)
            {
                if (double.TryParse(_valueAsString, out resultDouble))
                {
                    _valueAsDouble = resultDouble;
                }
                else if (decimal.TryParse(_valueAsString, out resultDecimal))
                {
                    _valueAsDouble = Convert.ToDouble(resultDecimal);
                }
                else if (long.TryParse(_valueAsString, out resultLong))
                {
                    _valueAsDouble = Convert.ToDouble(resultLong);
                }
                else if (_valueAsString.Equals(Tokens.True))
                {
                    _valueAsDouble = -1.0;
                }
                else if (_valueAsString.Equals(Tokens.False))
                {
                    _valueAsDouble = 0.0;
                }
            }
            return _valueAsDouble;
        }

        public bool? AsBoolean()
        {
            if (!_valueAsBoolean.HasValue)
            {
                if (_valueAsString.Equals(Tokens.True))
                {
                    _valueAsBoolean = true;
                }
                else if (_valueAsString.Equals(Tokens.False))
                {
                    _valueAsBoolean = false;
                }
                else if (long.TryParse(_valueAsString, out resultLong))
                {
                    _valueAsBoolean = resultLong != 0;
                }
                else if (double.TryParse(_valueAsString, out resultDouble))
                {
                    _valueAsBoolean = Math.Abs(resultDouble) > 0.00000001;
                }
                else if (decimal.TryParse(_valueAsString, out resultDecimal))
                {
                    _valueAsBoolean = Math.Abs(resultDecimal) > 0.0000001M;
                }
            }
            return _valueAsBoolean;
        }

        private long? SafeConvertToLong<T>(T value)
        {
            try
            {
                return Convert.ToInt64(value);
            }
            catch (OverflowException)
            {
                return null;
            }
        }

        private decimal? SafeConvertToDecimal<T>(T value)
        {
            try
            {
                return Convert.ToDecimal(value);
            }
            catch (OverflowException)
            {
                return null;
            }
        }
    }
#endregion
}
