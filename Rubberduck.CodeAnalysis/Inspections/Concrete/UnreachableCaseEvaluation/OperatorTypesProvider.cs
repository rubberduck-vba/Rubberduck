using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete.UnreachableCaseEvaluation
{
    internal class OperatorTypesProvider
    {
        private readonly (string lhs, string rhs) _operandTypeNames;
        private readonly string _opSymbol;
        private string _operatorDeclaredTypeName;
        private string _operatorEffectiveTypeName;

        private static OperatorTypesLookup _arithmeticDeclaredTypes;
        private static OperatorTypesLookup _logicalDeclaredTypes;
        private static OperatorTypesLookup _logicalEffectiveTypes;
        private static OperatorTypesLookup _relationalEffectiveTypes;
        private static OperatorTypesLookup _relationalDeclaredTypes;

        public OperatorTypesProvider((string lhsTypeName, string rhsTypeName) operands, string opSymbol)
        {
            _operandTypeNames = operands;
            _opSymbol = opSymbol;
            InitializeArithmeticDeclaredTypes(ref _arithmeticDeclaredTypes);
            InitializeLogicalDeclaredTypes(ref _logicalDeclaredTypes);
            InitializeLogicalEffectiveTypes(ref _logicalEffectiveTypes, _logicalDeclaredTypes);
            InitializeRelationalDeclaredTypes(ref _relationalDeclaredTypes);
            InitializeRelationalEffectiveTypes(ref _relationalEffectiveTypes);
        }

        public OperatorTypesProvider(string operandTypeName, string opSymbol)
            : this((operandTypeName, null), opSymbol) { }

        public bool IsMismatch
        {
            get
            {
                if (_operatorDeclaredTypeName != null)
                {
                    return false;
                }
                try
                {
                    ResolveTypes();
                    return false;
                }
                catch (KeyNotFoundException)
                {
                    return true;
                }
            }
        }

        public string OperatorDeclaredType => _operatorDeclaredTypeName ?? ResolveTypes().Item1;

        public string OperatorEffectiveType => _operatorEffectiveTypeName ?? ResolveTypes().Item2;

        //To support testing
        public string OperatorDeclaredTypeArithmeticDefault()
        {
            _operatorDeclaredTypeName = _arithmeticDeclaredTypes.TypeName(_operandTypeNames, _opSymbol);
            return _operatorDeclaredTypeName;
        }

        //To support testing
        public string OperatorDeclaredTypeLogicalDefault()
        {
            _operatorDeclaredTypeName = _logicalDeclaredTypes.TypeName(_operandTypeNames, _opSymbol);
            return _operatorDeclaredTypeName;
        }

        private (string,string) ResolveTypes()
        {
            _operatorDeclaredTypeName = null;
            if (ArithmeticOperators.Includes(_opSymbol))
            {
                _operatorDeclaredTypeName = _arithmeticDeclaredTypes.TypeName(_operandTypeNames, _opSymbol);
                _operatorEffectiveTypeName = _operatorDeclaredTypeName;
            }
            else if (RelationalOperators.Includes(_opSymbol))
            {
                _operatorDeclaredTypeName = _relationalDeclaredTypes.TypeName(_operandTypeNames, _opSymbol);
                _operatorEffectiveTypeName = _relationalEffectiveTypes.TypeName(_operandTypeNames, _opSymbol);
            }
            else if (LogicalOperators.Incudes(_opSymbol))
            {
                _operatorDeclaredTypeName = _logicalDeclaredTypes.TypeName(_operandTypeNames, _opSymbol);
                _operatorEffectiveTypeName = _logicalEffectiveTypes.TypeName(_operandTypeNames, _opSymbol);
            }
            if (_operatorDeclaredTypeName is null)
            {
                throw new KeyNotFoundException($"Unhandled operation symbol: {_opSymbol}");
            }
            return (_operatorDeclaredTypeName, _operatorEffectiveTypeName);
        }

        private static string ArithmeticDeclaredUnary(string typeName, string opSymbol)
        {
            var resultTypeName = typeName;
            if (typeName.Equals(Tokens.Boolean))
            {
                resultTypeName = Tokens.Integer;
            }
            if (typeName.Equals(Tokens.String))
            {
                resultTypeName = Tokens.Double;
            }
            if (opSymbol is null || opSymbol.Equals(string.Empty))
            {
                return resultTypeName;
            }
            else if (opSymbol.Equals(ArithmeticOperators.MINUS) && typeName.Equals(Tokens.Byte))
            {
                return Tokens.Integer;
            }
            return resultTypeName;
        }

        private static string LogicalDeclaredUnary(string typeName, string opSymbol)
        {
            var unaryLong = FloatingPointAndFixedPointNumericTypesStringLongAndDate();
            if (unaryLong.Contains(typeName))
            {
                return Tokens.Long;
            }
            return typeName;
        }

        private static string LogicalEffectiveUnary(string typeName, string opSymbol)
        {
            if (typeName.Equals(Tokens.Boolean))
            {
                return Tokens.Integer;
            }
            if (typeName.Equals(Tokens.String))
            {
                return Tokens.Double;
            }
            return typeName;
        }

        private static string RelationalType((string lhs,string rhs ) operands, string opSymbol)
        {
            if (!operands.lhs.Equals(Tokens.Variant) || !operands.rhs.Equals(Tokens.Variant))
            {
                return Tokens.Boolean;
            }
            return Tokens.Variant;
        }

        private static void InitializeArithmeticDeclaredTypes(ref OperatorTypesLookup typeProvider)
        {
            if (!typeProvider.IsEmpty)
            {
                return;
            }

            typeProvider.UnaryResolver = ArithmeticDeclaredUnary;

            typeProvider.Add(new string[] { Tokens.Byte }, new string[] { Tokens.Byte }, Tokens.Byte);

            typeProvider.Add(new string[] { Tokens.Boolean, Tokens.Integer }, new string[] { Tokens.Boolean, Tokens.Integer, Tokens.Byte }, Tokens.Integer);
            typeProvider.Add(new string[] { Tokens.Boolean, Tokens.Integer, Tokens.Byte }, new string[] { Tokens.Boolean, Tokens.Integer }, Tokens.Integer);

            typeProvider.Add(new string[] { Tokens.Long }, IntegralNumericTypesExceptLongLong(), Tokens.Long);
            typeProvider.Add(IntegralNumericTypesExceptLongLong(), new string[] { Tokens.Long }, Tokens.Long);

            typeProvider.Add(new string[] { Tokens.LongLong }, IntegralNumericTypes, Tokens.LongLong);
            typeProvider.Add(IntegralNumericTypes, new string[] { Tokens.LongLong }, Tokens.LongLong);

            typeProvider.Add(new string[] { Tokens.Single }, new string[] { Tokens.Single, Tokens.Boolean, Tokens.Integer, Tokens.Byte }, Tokens.Single);
            typeProvider.Add(new string[] { Tokens.Single, Tokens.Boolean, Tokens.Integer, Tokens.Byte }, new string[] { Tokens.Single }, Tokens.Single);

            typeProvider.Add(new string[] { Tokens.Single }, new string[] { Tokens.Long, Tokens.LongLong }, Tokens.Double);
            typeProvider.Add(new string[] { Tokens.Long, Tokens.LongLong }, new string[] { Tokens.Single }, Tokens.Double);

            typeProvider.Add(new string[] { Tokens.Double, Tokens.String }, IntegralOrFloatingPointNumericTypesAndString(), Tokens.Double);
            typeProvider.Add(IntegralOrFloatingPointNumericTypesAndString(), new string[] { Tokens.Double, Tokens.String }, Tokens.Double);

            typeProvider.Add(new string[] { Tokens.Currency }, AllNumericTypesAndString(), Tokens.Currency);
            typeProvider.Add(AllNumericTypesAndString(), new string[] { Tokens.Currency }, Tokens.Currency);

            typeProvider.Add(new string[] { Tokens.Date }, AllNumericTypesAndStringAndDate(), Tokens.Date);
            typeProvider.Add(AllNumericTypesAndStringAndDate(), new string[] { Tokens.Date }, Tokens.Date);

            typeProvider.Add(new string[] { Tokens.Variant }, AllNumericTypesAndStringAndDate(), Tokens.Variant);
            typeProvider.Add(AllNumericTypesAndStringAndDate(), new string[] { Tokens.Variant }, Tokens.Variant);

            typeProvider.AddOperatorException(ArithmeticOperators.PLUS, new string[] { Tokens.String }, new string[] { Tokens.String }, Tokens.String);

            typeProvider.AddOperatorException(ArithmeticOperators.MODULO, new string[] { Tokens.Date, Tokens.Currency, Tokens.Double, Tokens.Single, Tokens.String }, new string[] { Tokens.Date, Tokens.Currency, Tokens.Double, Tokens.Single, Tokens.LongLong, Tokens.Long, Tokens.Boolean, Tokens.Integer, Tokens.Byte, Tokens.String }, Tokens.Long);
            typeProvider.AddOperatorException(ArithmeticOperators.MODULO, new string[] { Tokens.Date, Tokens.Currency, Tokens.Double, Tokens.Single, Tokens.LongLong, Tokens.Long, Tokens.Boolean, Tokens.Integer, Tokens.Byte, Tokens.String }, new string[] { Tokens.Date, Tokens.Currency, Tokens.Double, Tokens.Single, Tokens.String }, Tokens.Long);

            typeProvider.AddOperatorException(ArithmeticOperators.INTEGER_DIVIDE, new string[] { Tokens.Date, Tokens.Currency, Tokens.Double, Tokens.Single, Tokens.String }, new string[] { Tokens.Date, Tokens.Currency, Tokens.Double, Tokens.Single, Tokens.LongLong, Tokens.Long, Tokens.Boolean, Tokens.Integer, Tokens.Byte, Tokens.String }, Tokens.Long);
            typeProvider.AddOperatorException(ArithmeticOperators.INTEGER_DIVIDE, new string[] { Tokens.Date, Tokens.Currency, Tokens.Double, Tokens.Single, Tokens.LongLong, Tokens.Long, Tokens.Boolean, Tokens.Integer, Tokens.Byte, Tokens.String }, new string[] { Tokens.Date, Tokens.Currency, Tokens.Double, Tokens.Single, Tokens.String }, Tokens.Long);

            typeProvider.AddOperatorException(ArithmeticOperators.MULTIPLY, new string[] { Tokens.Currency }, new string[] { Tokens.Double, Tokens.Single, Tokens.String }, Tokens.Double);
            typeProvider.AddOperatorException(ArithmeticOperators.MULTIPLY, new string[] { Tokens.Double, Tokens.Single, Tokens.String }, new string[] { Tokens.Currency }, Tokens.Double);

            typeProvider.AddOperatorException(ArithmeticOperators.DIVIDE, new string[] { Tokens.LongLong, Tokens.Long, Tokens.Boolean, Tokens.Integer, Tokens.Byte }, new string[] { Tokens.LongLong, Tokens.Long, Tokens.Boolean, Tokens.Integer, Tokens.Byte }, Tokens.Double);
            typeProvider.AddOperatorException(ArithmeticOperators.DIVIDE, new string[] { Tokens.Date, Tokens.Currency, Tokens.Double, Tokens.String }, new string[] { Tokens.Date, Tokens.Currency, Tokens.Double, Tokens.Single, Tokens.LongLong, Tokens.Long, Tokens.Boolean, Tokens.Integer, Tokens.Byte, Tokens.String }, Tokens.Double);
            typeProvider.AddOperatorException(ArithmeticOperators.DIVIDE, new string[] { Tokens.Date, Tokens.Currency, Tokens.Double, Tokens.Single, Tokens.LongLong, Tokens.Long, Tokens.Boolean, Tokens.Integer, Tokens.Byte, Tokens.String }, new string[] { Tokens.Date, Tokens.Currency, Tokens.Double, Tokens.String }, Tokens.Double);

            typeProvider.AddOperatorException(ArithmeticOperators.MINUS, new string[] { Tokens.Date }, new string[] { Tokens.Date }, Tokens.Double);

            typeProvider.AddOperatorException(ArithmeticOperators.EXPONENT, AllTypesExceptArraysAndUDTs(), AllTypesExceptArraysAndUDTs(), Tokens.Double);

            typeProvider.AddOperatorException(ArithmeticOperators.AMPERSAND, AllNumericTypesAndStringAndDate(), AllNumericTypesAndStringAndDate(), Tokens.String);

            typeProvider.AddOperatorException(ArithmeticOperators.AMPERSAND, AllTypesExceptArraysAndUDTs(), new string[] { Tokens.Variant }, Tokens.Variant);
            typeProvider.AddOperatorException(ArithmeticOperators.AMPERSAND, new string[] { Tokens.Variant }, AllTypesExceptArraysAndUDTs(), Tokens.Variant);
        }

        private static void InitializeLogicalDeclaredTypes(ref OperatorTypesLookup typeProvider)
        {
            if (!typeProvider.IsEmpty)
            {
                return;
            }

            typeProvider.UnaryResolver = LogicalDeclaredUnary;

            typeProvider.Add(new List<string>() { Tokens.Byte }, new List<string>() { Tokens.Byte }, Tokens.Byte);
            typeProvider.Add(new List<string>() { Tokens.Boolean }, new List<string>() { Tokens.Boolean }, Tokens.Boolean);

            typeProvider.Add(new List<string>() { Tokens.Byte, Tokens.Integer }, new List<string>() { Tokens.Boolean, Tokens.Integer, Tokens.Byte }, Tokens.Integer);
            typeProvider.Add(new List<string>() { Tokens.Boolean, Tokens.Integer }, new List<string>() { Tokens.Byte, Tokens.Integer }, Tokens.Integer);

            typeProvider.Add(FloatingPointAndFixedPointNumericTypesStringLongAndDate(), AllNumericTypesExceptLongLongStringAndDate(), Tokens.Long);
            typeProvider.Add(AllNumericTypesExceptLongLongStringAndDate(), FloatingPointAndFixedPointNumericTypesStringLongAndDate(), Tokens.Long);

            typeProvider.Add(new List<string>() { Tokens.LongLong }, AllNumericTypesStringAndDate(), Tokens.LongLong);
            typeProvider.Add(AllNumericTypesStringAndDate(), new List<string>() { Tokens.LongLong }, Tokens.LongLong);

            typeProvider.Add(AllTypesExceptArraysAndUDTs(), new List<string>() { Tokens.Variant }, Tokens.Variant);
            typeProvider.Add(new List<string>() { Tokens.Variant }, AllTypesExceptArraysAndUDTs(), Tokens.Variant);
        }

        private static void InitializeLogicalEffectiveTypes(ref OperatorTypesLookup typeProvider, OperatorTypesLookup logicalDeclaredTypes)
        {
            typeProvider.UnaryResolver = LogicalEffectiveUnary;
            typeProvider.BinaryResolver = logicalDeclaredTypes.TypeName;
        }

        private static void InitializeRelationalDeclaredTypes(ref OperatorTypesLookup typeProvider)
        {
            typeProvider.BinaryResolver = RelationalType;
        }

        private static void InitializeRelationalEffectiveTypes(ref OperatorTypesLookup typeProvider)
        {
            if (!typeProvider.IsEmpty)
            {
                return;
            }

            typeProvider.Add(new List<string>() { Tokens.Byte }, new List<string>() { Tokens.Byte, Tokens.String }, Tokens.Byte);
            typeProvider.Add(new List<string>() { Tokens.Byte, Tokens.String }, new List<string>() { Tokens.Byte }, Tokens.Byte);

            typeProvider.Add(new List<string>() { Tokens.Boolean }, new List<string>() { Tokens.Boolean, Tokens.String }, Tokens.Boolean);
            typeProvider.Add(new List<string>() { Tokens.Boolean, Tokens.String }, new List<string>() { Tokens.Boolean }, Tokens.Boolean);

            typeProvider.Add(new List<string>() { Tokens.Integer }, new List<string>() { Tokens.Byte, Tokens.Boolean, Tokens.Integer, Tokens.String }, Tokens.Integer);
            typeProvider.Add(new List<string>() { Tokens.Byte, Tokens.Boolean, Tokens.Integer, Tokens.String }, new List<string>() { Tokens.Integer }, Tokens.Integer);

            typeProvider.Add(new List<string>() { Tokens.Boolean }, new List<string>() { Tokens.Byte }, Tokens.Integer);
            typeProvider.Add(new List<string>() { Tokens.Byte }, new List<string>() { Tokens.Boolean }, Tokens.Integer);

            typeProvider.Add(new List<string>() { Tokens.Long }, new List<string>() { Tokens.Byte, Tokens.Boolean, Tokens.Integer, Tokens.Long, Tokens.String }, Tokens.Long);
            typeProvider.Add(new List<string>() { Tokens.Byte, Tokens.Boolean, Tokens.Integer, Tokens.Long, Tokens.String }, new List<string>() { Tokens.Long }, Tokens.Long);

            typeProvider.Add(new List<string>() { Tokens.LongLong }, AllNumericTypesExceptLongLongStringAndDate(), Tokens.LongLong);
            typeProvider.Add(AllNumericTypesExceptLongLongStringAndDate(), new List<string>() { Tokens.LongLong }, Tokens.LongLong);

            typeProvider.Add(new List<string>() { Tokens.Single }, new List<string>() { Tokens.Byte, Tokens.Boolean, Tokens.Integer, Tokens.Single, Tokens.Double, Tokens.String }, Tokens.Single);
            typeProvider.Add(new List<string>() { Tokens.Byte, Tokens.Boolean, Tokens.Integer, Tokens.Single, Tokens.Double, Tokens.String }, new List<string>() { Tokens.Single }, Tokens.Single);

            typeProvider.Add(new List<string>() { Tokens.Single }, new List<string>() { Tokens.Long }, Tokens.Double);
            typeProvider.Add(new List<string>() { Tokens.Long }, new List<string>() { Tokens.Single }, Tokens.Double);

            var listing = AllNumericTypesAndString();
            listing.Add(Tokens.Double);
            typeProvider.Add(new List<string>() { Tokens.Double }, listing, Tokens.Double);
            typeProvider.Add(listing, new List<string>() { Tokens.Double }, Tokens.Double);

            typeProvider.Add(new List<string>() { Tokens.String }, new List<string>() { Tokens.String }, Tokens.String);

            typeProvider.Add(new List<string>() { Tokens.Currency }, AllNumericTypesAndString(), Tokens.Currency);
            typeProvider.Add(AllNumericTypesAndString(), new List<string>() { Tokens.Currency }, Tokens.Currency);

            typeProvider.Add(new List<string>() { Tokens.Date }, AllNumericTypesAndStringAndDate(), Tokens.Date);
            typeProvider.Add(AllNumericTypesAndStringAndDate(), new List<string>() { Tokens.Date }, Tokens.Date);
        }

        private struct OperatorTypesLookup
        {
            private List<(IEnumerable<string>, IEnumerable<string>, string)> _defaults;
            private Dictionary<string, OperatorTypesLookup> _exceptions;

            public Func<(string,string), string, string> BinaryResolver;

            public Func<string, string, string> UnaryResolver;

            public void Add(IEnumerable<string> leftTypes, IEnumerable<string> rightTypes, string operatorDeclaredType)
            {
                if (_defaults is null)
                {
                    _defaults = new List<(IEnumerable<string>, IEnumerable<string>, string)>();
                }
                _defaults.Add((leftTypes, rightTypes, operatorDeclaredType));
            }

            public void AddOperatorException(string opSymbol, IEnumerable<string> leftTypes, IEnumerable<string> rightTypes, string operatorDeclaredType)
            {
                if (_exceptions is null)
                {
                    _exceptions = new Dictionary<string, OperatorTypesLookup>();
                }

                var operatorException = new OperatorTypesLookup();
                if (_exceptions.ContainsKey(opSymbol))
                {
                    operatorException = _exceptions[opSymbol];
                }
                operatorException.Add(leftTypes, rightTypes, operatorDeclaredType);
                _exceptions[opSymbol] = operatorException;
            }

            public bool IsEmpty => _defaults is null || _defaults.Count == 0;

            public string TypeName((string lhsTypeName, string rhsTypeName) operands, string opSymbol)
            {
                var defaultType = DefaultType(operands);
                if (!IsUnaryOperand(operands))
                {
                    if (opSymbol != null && _exceptions != null && _exceptions.ContainsKey(opSymbol))
                    {
                        var exceptionType = _exceptions[opSymbol].TypeName(operands, opSymbol);
                        return exceptionType ?? defaultType;
                    }
                }
                return defaultType;
            }

            private string DefaultType((string lhsTypeName, string rhsTypeName) operands)
            {
                if (IsUnaryOperand(operands))
                {
                    if (UnaryResolver is null) { return null; }

                    return UnaryResolver(operands.lhsTypeName, string.Empty);
                }

                if (BinaryResolver != null)
                {
                    return BinaryResolver(operands, null);
                }

                foreach (var (lhsTypes, rhsTypes, opTypeName) in _defaults)
                {
                    if (lhsTypes.Contains(operands.lhsTypeName) && rhsTypes.Contains(operands.rhsTypeName))
                    {
                        return opTypeName;
                    }
                }
                return null;
            }

            private bool IsUnaryOperand((string lhsTypeName, string rhsTypeName) operands)
                => operands.rhsTypeName is null || operands.rhsTypeName.Equals(string.Empty);
        }

#region DefinedTypeLists
        public static List<string> IntegralNumericTypes = new List<string>()
        {
            Tokens.Byte,
            Tokens.Boolean,
            Tokens.Integer,
            Tokens.Long,
            Tokens.LongLong,
        };

        private static List<string> FloatingPointAndFixedPointNumericTypes = new List<string>()
        {
            Tokens.Single,
            Tokens.Double,
            Tokens.Currency,
        };

        private static List<string> IntegralNumericTypesExceptLongLong()
        {
            var results = new List<string>();
            results.AddRange(IntegralNumericTypes);
            results.Remove(Tokens.LongLong);
            return results;
        }

        private static List<string> FloatingPointAndFixedPointNumericTypesStringLongAndDate()
        {
            var results = new List<string>();
            results.AddRange(FloatingPointAndFixedPointNumericTypes);
            results.Add(Tokens.String);
            results.Add(Tokens.Long);
            results.Add(Tokens.Date);
            return results;
        }

        private static List<string> AllNumericTypes()
        {
            var results = new List<string>();
            results.AddRange(IntegralNumericTypes);
            results.AddRange(FloatingPointAndFixedPointNumericTypes);
            return results;
        }

        private static List<string> AllNumericTypesExceptLongLongStringAndDate()
        {
            var results = new List<string>();
            results.AddRange(IntegralNumericTypes);
            results.AddRange(FloatingPointAndFixedPointNumericTypes);
            results.Remove(Tokens.LongLong);
            results.Remove(Tokens.String);
            results.Remove(Tokens.Date);
            return results;
        }

        private static List<string> AllNumericTypesStringAndDate()
        {
            var results = new List<string>();
            results.AddRange(IntegralNumericTypes);
            results.AddRange(FloatingPointAndFixedPointNumericTypes);
            results.Add(Tokens.String);
            results.Add(Tokens.Date);
            return results;
        }

        private static List<string> IntegralOrFloatingPointNumericTypesAndString()
        {
            var results = new List<string>();
            results.AddRange(IntegralNumericTypes);
            results.AddRange(FloatingPointAndFixedPointNumericTypes);
            results.Add(Tokens.String);
            results.Remove(Tokens.Currency);
            return results;
        }

        private static List<string> AllNumericTypesAndString()
        {
            var results = new List<string>();
            results.AddRange(IntegralNumericTypes);
            results.AddRange(FloatingPointAndFixedPointNumericTypes);
            results.Add(Tokens.String);
            return results;
        }

        private static List<string> AllNumericTypesAndStringAndDate()
        {
            var results = new List<string>();
            results.AddRange(IntegralNumericTypes);
            results.AddRange(FloatingPointAndFixedPointNumericTypes);
            results.Add(Tokens.String);
            results.Add(Tokens.Date);
            return results;
        }

        private static List<string> AllTypesExceptArraysAndUDTs()
        {
            var results = new List<string>();
            results.AddRange(IntegralNumericTypes);
            results.AddRange(FloatingPointAndFixedPointNumericTypes);
            results.Add(Tokens.String);
            results.Add(Tokens.Date);
            results.Add(Tokens.Variant);
            return results;
        }
    }
    #endregion
}
