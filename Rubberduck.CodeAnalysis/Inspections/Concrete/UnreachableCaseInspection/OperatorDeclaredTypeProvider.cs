using Rubberduck.Parsing.Grammar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IBinaryOperatorDeclaredType
    {
        string BinaryOperatorDeclaredType((string lhsTypeName, string rhsTypeName) operands, string opSymbol);
    }

    public interface IUnaryOperatorDeclaredType
    {
        string UnaryOperatorDeclaredType(string operandTypeName, string opSymbol);
    }

    public enum OperatorDeclaredTypeProviderTypes
    {
        Arithmetic = 1,
        Relational,
        Logical
    };

    public class OperatorDeclaredTypeProvider
    {
        private readonly string _lhsTypeName;
        private readonly string _rhsTypeName;
        private readonly string _opSymbol;
        private readonly int _opType;
        private readonly bool _isBinary;
        private readonly string _operatorDeclaredTypeName;

        public OperatorDeclaredTypeProvider((string lhsTypeName, string rhsTypeName) operands, OperatorDeclaredTypeProviderTypes opType)
        {
            _lhsTypeName = operands.lhsTypeName;
            _rhsTypeName = operands.rhsTypeName;
            _opSymbol = null;
            _isBinary = true;
            _opType = Convert.ToInt32(opType);
            _operatorDeclaredTypeName = ResolveDeclaredType();
        }

        public OperatorDeclaredTypeProvider((string lhsTypeName, string rhsTypeName) operands, string opSymbol)
        {
            _lhsTypeName = operands.lhsTypeName;
            _rhsTypeName = operands.rhsTypeName;
            _opSymbol = opSymbol;
            _isBinary = true;
            _opType = 0;
            _operatorDeclaredTypeName = ResolveDeclaredType();
        }

        public OperatorDeclaredTypeProvider(string typeName, string opSymbol)
        {
            _lhsTypeName = typeName;
            _opSymbol = opSymbol;
            _isBinary = false;
            _opType = 0;
            _operatorDeclaredTypeName = ResolveDeclaredType();
        }

        public string OperatorDeclaredType => _operatorDeclaredTypeName;

        public static List<string> IntegralNumericTypes = new List<string>()
            {
                Tokens.Byte,
                Tokens.Boolean,
                Tokens.Integer,
                Tokens.Long,
                Tokens.LongLong,
            };

        private string ResolveDeclaredType()
        {
            if (ArithmeticOperators.Includes(_opSymbol) || _opType == Convert.ToInt64(OperatorDeclaredTypeProviderTypes.Arithmetic))
            {
                if (_isBinary)
                {
                    return new OperatorDeclaredTypeArithmetic().BinaryOperatorDeclaredType((_lhsTypeName, _rhsTypeName), _opSymbol);
                }
                return new OperatorDeclaredTypeArithmetic().UnaryOperatorDeclaredType(_lhsTypeName, _opSymbol);
            }

            if (RelationalOperators.Includes(_opSymbol) || _opType == Convert.ToInt64(OperatorDeclaredTypeProviderTypes.Relational))
                {
                return new OperatorDeclaredTypeRelational().BinaryOperatorDeclaredType((_lhsTypeName, _rhsTypeName), _opSymbol);
            }

            if (LogicalOperators.Incudes(_opSymbol) || _opType == Convert.ToInt64(OperatorDeclaredTypeProviderTypes.Logical))
            {
                if (_isBinary)
                {
                    return new OperatorDeclaredTypeLogical().BinaryOperatorDeclaredType((_lhsTypeName, _rhsTypeName), _opSymbol);
                }
                return new OperatorDeclaredTypeLogical().UnaryOperatorDeclaredType(_lhsTypeName, _opSymbol);
            }

            return string.Empty;
        }

        private class OperatorDeclaredTypeBase
        {

            private List<(List<string> lhsTypes, List<string> rhsTypes, string typeName)> OperatorDeclaredTypesBinaryArithmetic { set; get; } = null;

            private List<(List<string> lhsTypes, List<string> rhsTypes, string typeName)> OperatorDeclaredTypesBinaryLogical { set; get; } = null;

            private static List<string> FloatingPointAndFixedPointNumericTypes = new List<string>()
            {
                Tokens.Single,
                Tokens.Double,
                Tokens.Currency,
            };

            protected static List<string> IntegralNumericTypesExceptLongLong()
            {
                var results = new List<string>();
                results.AddRange(IntegralNumericTypes);
                results.Remove(Tokens.LongLong);
                return results;
            }

            protected static List<string> FloatingPointAndFixedPointNumericTypesStringLongAndDate()
            {
                var results = new List<string>();
                results.AddRange(FloatingPointAndFixedPointNumericTypes);
                results.Add(Tokens.String);
                results.Add(Tokens.Long);
                results.Add(Tokens.Date);
                return results;
            }

            protected static List<string> AllNumericTypes()
            {
                var results = new List<string>();
                results.AddRange(IntegralNumericTypes);
                results.AddRange(FloatingPointAndFixedPointNumericTypes);
                return results;
            }

            protected static List<string> AllNumericTypesExceptLongLongStringAndDate()
            {
                var results = new List<string>();
                results.AddRange(IntegralNumericTypes);
                results.AddRange(FloatingPointAndFixedPointNumericTypes);
                results.Remove(Tokens.LongLong);
                results.Remove(Tokens.String);
                results.Remove(Tokens.Date);
                return results;
            }

            protected static List<string> AllNumericTypesStringAndDate()
            {
                var results = new List<string>();
                results.AddRange(IntegralNumericTypes);
                results.AddRange(FloatingPointAndFixedPointNumericTypes);
                results.Add(Tokens.String);
                results.Add(Tokens.Date);
                return results;
            }

            protected static List<string> IntegralOrFloatingPointNumericTypesAndString() 
            {
                var results = new List<string>();
                results.AddRange(IntegralNumericTypes);
                results.AddRange(FloatingPointAndFixedPointNumericTypes);
                results.Add(Tokens.String);
                results.Remove(Tokens.Currency);
                return results;
            }

            protected static List<string> AllNumericTypesAndString()
            {
                var results = new List<string>();
                results.AddRange(IntegralNumericTypes);
                results.AddRange(FloatingPointAndFixedPointNumericTypes);
                results.Add(Tokens.String);
                return results;
            }

            protected static List<string> AllNumericTypesAndStringAndDate()
            {
                var results = new List<string>();
                results.AddRange(IntegralNumericTypes);
                results.AddRange(FloatingPointAndFixedPointNumericTypes);
                results.Add(Tokens.String);
                results.Add(Tokens.Date);
                return results;
            }

            protected static List<string> AllTypesExceptArraysAndUDTs()
            {
                var results = new List<string>();
                results.AddRange(IntegralNumericTypes);
                results.AddRange(FloatingPointAndFixedPointNumericTypes);
                results.Add(Tokens.String);
                results.Add(Tokens.Date);
                results.Add(Tokens.Variant);
                return results;
            }

            public string OperatorDeclaredTypeDefaultArithmeticBinary(string lhsTypeName, string rhsTypeName)
            {
                if (OperatorDeclaredTypesBinaryArithmetic is null)
                {
                    OperatorDeclaredTypesBinaryArithmetic = new List<(List<string> lhsTypes, List<string> rhsTypes, string typeName)>();

                    OperatorDeclaredTypesBinaryArithmetic.Add(((new List<string>() { Tokens.Byte }, new List<string>() { Tokens.Byte }, Tokens.Byte)));

                    OperatorDeclaredTypesBinaryArithmetic.Add(((new List<string>() { Tokens.Boolean, Tokens.Integer }, new List<string>() { Tokens.Boolean, Tokens.Integer, Tokens.Byte }, Tokens.Integer)));
                    OperatorDeclaredTypesBinaryArithmetic.Add(((new List<string>() { Tokens.Boolean, Tokens.Integer, Tokens.Byte }, new List<string>() { Tokens.Boolean, Tokens.Integer }, Tokens.Integer)));

                    OperatorDeclaredTypesBinaryArithmetic.Add(((new List<string>() { Tokens.Long }, IntegralNumericTypesExceptLongLong(), Tokens.Long)));
                    OperatorDeclaredTypesBinaryArithmetic.Add(((IntegralNumericTypesExceptLongLong(), new List<string>() { Tokens.Long }, Tokens.Long)));

                    OperatorDeclaredTypesBinaryArithmetic.Add(((new List<string>() { Tokens.LongLong }, IntegralNumericTypes, Tokens.LongLong)));
                    OperatorDeclaredTypesBinaryArithmetic.Add(((IntegralNumericTypes, new List<string>() { Tokens.LongLong }, Tokens.LongLong)));

                    OperatorDeclaredTypesBinaryArithmetic.Add(((new List<string>() { Tokens.Single }, new List<string>() { Tokens.Single, Tokens.Boolean, Tokens.Integer, Tokens.Byte }, Tokens.Single)));
                    OperatorDeclaredTypesBinaryArithmetic.Add(((new List<string>() { Tokens.Single, Tokens.Boolean, Tokens.Integer, Tokens.Byte }, new List<string>() { Tokens.Single }, Tokens.Single)));

                    OperatorDeclaredTypesBinaryArithmetic.Add(((new List<string>() { Tokens.Single }, new List<string>() { Tokens.Long, Tokens.LongLong }, Tokens.Double)));
                    OperatorDeclaredTypesBinaryArithmetic.Add(((new List<string>() { Tokens.Long, Tokens.LongLong }, new List<string>() { Tokens.Single }, Tokens.Double)));

                    OperatorDeclaredTypesBinaryArithmetic.Add(((new List<string>() { Tokens.Double, Tokens.String }, IntegralOrFloatingPointNumericTypesAndString(), Tokens.Double)));
                    OperatorDeclaredTypesBinaryArithmetic.Add(((IntegralOrFloatingPointNumericTypesAndString(), new List<string>() { Tokens.Double, Tokens.String }, Tokens.Double)));

                    OperatorDeclaredTypesBinaryArithmetic.Add(((new List<string>() { Tokens.Currency }, AllNumericTypesAndString(), Tokens.Currency)));
                    OperatorDeclaredTypesBinaryArithmetic.Add(((AllNumericTypesAndString(), new List<string>() { Tokens.Currency }, Tokens.Currency)));

                    OperatorDeclaredTypesBinaryArithmetic.Add(((new List<string>() { Tokens.Date }, AllNumericTypesAndStringAndDate(), Tokens.Date)));
                    OperatorDeclaredTypesBinaryArithmetic.Add(((AllNumericTypesAndStringAndDate(), new List<string>() { Tokens.Date }, Tokens.Date)));
                }

                var result = RetrieveOperatorDeclaredTypeName(OperatorDeclaredTypesBinaryArithmetic, lhsTypeName, rhsTypeName);
                return result;
            }

            public string OperatorDeclaredTypeDefaultArithmeticUnary(string typeName)
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

            public string OperatorDeclaredTypeDefaultLogical(string lhsTypeName, string rhsTypeName)
            {
                if (OperatorDeclaredTypesBinaryLogical is null)
                {
                    OperatorDeclaredTypesBinaryLogical = new List<(List<string> lhsTypes, List<string> rhsTypes, string typeName)>();

                    OperatorDeclaredTypesBinaryLogical.Add(((new List<string>() { Tokens.Byte }, new List<string>() { Tokens.Byte }, Tokens.Byte)));
                    OperatorDeclaredTypesBinaryLogical.Add(((new List<string>() { Tokens.Boolean }, new List<string>() { Tokens.Boolean }, Tokens.Boolean)));

                    OperatorDeclaredTypesBinaryLogical.Add(((new List<string>() { Tokens.Byte, Tokens.Integer }, new List<string>() { Tokens.Boolean, Tokens.Integer, Tokens.Byte }, Tokens.Integer)));
                    OperatorDeclaredTypesBinaryLogical.Add(((new List<string>() { Tokens.Boolean, Tokens.Integer }, new List<string>() { Tokens.Byte, Tokens.Integer }, Tokens.Integer)));

                    OperatorDeclaredTypesBinaryLogical.Add(((FloatingPointAndFixedPointNumericTypesStringLongAndDate(), AllNumericTypesExceptLongLongStringAndDate(), Tokens.Long)));
                    OperatorDeclaredTypesBinaryLogical.Add(((AllNumericTypesExceptLongLongStringAndDate(), FloatingPointAndFixedPointNumericTypesStringLongAndDate(), Tokens.Long)));

                    OperatorDeclaredTypesBinaryLogical.Add(((new List<string>() { Tokens.LongLong }, AllNumericTypesStringAndDate(), Tokens.LongLong)));
                    OperatorDeclaredTypesBinaryLogical.Add(((AllNumericTypesStringAndDate(), new List<string>() { Tokens.LongLong }, Tokens.LongLong)));

                    OperatorDeclaredTypesBinaryLogical.Add(((AllTypesExceptArraysAndUDTs(), new List<string>() { Tokens.Variant }, Tokens.Variant)));
                }

                var result = RetrieveOperatorDeclaredTypeName(OperatorDeclaredTypesBinaryLogical, lhsTypeName, rhsTypeName);
                return result;
            }

            protected string RetrieveOperatorDeclaredTypeName(List<(List<string> lhsTypes, List<string> rhsTypes, string typeName)> values, string lhsTypeName, string rhsTypeName)
            {
                foreach (var (lhsTypes, rhsTypes, typeName) in values)
                {
                    if (lhsTypes.Contains(lhsTypeName) && rhsTypes.Contains(rhsTypeName))
                    {
                        return typeName;
                    }
                }
                return string.Empty;
            }
        }

        private class OperatorDeclaredTypeArithmetic : OperatorDeclaredTypeBase, IBinaryOperatorDeclaredType, IUnaryOperatorDeclaredType
        {
            private static List<(List<string> lhsTypes, List<string> rhsTypes, string typeName)> OperatorDeclaredTypesMod { set; get; } = null;
            private static List<(List<string> lhsTypes, List<string> rhsTypes, string typeName)> OperatorDeclaredTypesMult { set; get; } = null;
            private static List<(List<string> lhsTypes, List<string> rhsTypes, string typeName)> OperatorDeclaredTypesDiv { set; get; } = null;

            public string UnaryOperatorDeclaredType(string operandTypeName, string opSymbol)
            {
                return OperatorDeclaredTypeDefaultArithmeticUnary(operandTypeName);
            }

            public string BinaryOperatorDeclaredType((string lhsTypeName, string rhsTypeName) operands, string opSymbol)
            {
                if (opSymbol != null)
                {
                    if (opSymbol.Equals(ArithmeticOperators.MULTIPLY))
                    {
                        return DetermineOperatorDeclaredTypeBinary(operands.lhsTypeName, operands.rhsTypeName, OperatorDeclaredTypeMultiply);
                    }
                    if (opSymbol.Equals(ArithmeticOperators.DIVIDE))
                    {
                        return DetermineOperatorDeclaredTypeBinary(operands.lhsTypeName, operands.rhsTypeName, OperatorDeclaredTypeDivide);
                    }
                    if (opSymbol.Equals(ArithmeticOperators.MODULO))
                    {
                        return DetermineOperatorDeclaredTypeBinary(operands.lhsTypeName, operands.rhsTypeName, OperatorDeclaredTypeMod);
                    }
                    if (opSymbol.Equals(ArithmeticOperators.INTEGER_DIVIDE))
                    {
                        return DetermineOperatorDeclaredTypeBinary(operands.lhsTypeName, operands.rhsTypeName, OperatorDeclaredTypeIntegerDivide);
                    }
                }

                return OperatorDeclaredTypeDefaultArithmeticBinary(operands.lhsTypeName, operands.rhsTypeName);
            }

            private string DetermineOperatorDeclaredTypeBinary(string lhsTypeName, string rhsTypeName, Func<string, string, string> ExceptionType = null)
            {
                var operatorDeclaredType = OperatorDeclaredTypeDefaultArithmeticBinary(lhsTypeName, rhsTypeName);
                if (ExceptionType != null)
                {
                    var exceptionType = ExceptionType(lhsTypeName, rhsTypeName);
                    operatorDeclaredType = exceptionType.Equals(string.Empty) ? operatorDeclaredType : exceptionType;
                }
                return operatorDeclaredType;
            }

            private string OperatorDeclaredTypeMultiply(string lhsTypeName, string rhsTypeName)
            {
                //Apply '*' Operator exceptions to the default OperatorDeclaredTypesBinary per VBA spec 5.6.9.3
                if (OperatorDeclaredTypesMult is null)
                {
                    OperatorDeclaredTypesMult = new List<(List<string> lhsTypes, List<string> rhsTypes, string typeName)>();

                    OperatorDeclaredTypesMult.Add(((new List<string>() { Tokens.Currency }, new List<string>() { Tokens.Double, Tokens.Single, Tokens.String }, Tokens.Double)));
                    OperatorDeclaredTypesMult.Add(((new List<string>() { Tokens.Double, Tokens.Single, Tokens.String }, new List<string>() { Tokens.Currency }, Tokens.Double)));
                }

                return RetrieveOperatorDeclaredTypeName(OperatorDeclaredTypesMult, lhsTypeName, rhsTypeName);
            }

            private string OperatorDeclaredTypeDivide(string lhsTypeName, string rhsTypeName)
            {
                //Apply '/' Operator exceptions to the default OperatorDeclaredTypesBinary per VBA spec 5.6.9.3
                if (OperatorDeclaredTypesDiv is null)
                {
                    OperatorDeclaredTypesDiv = new List<(List<string> lhsTypes, List<string> rhsTypes, string typeName)>();

                    OperatorDeclaredTypesDiv.Add(((new List<string>() { Tokens.LongLong, Tokens.Long, Tokens.Boolean, Tokens.Integer, Tokens.Byte }, new List<string>() { Tokens.LongLong, Tokens.Long, Tokens.Boolean, Tokens.Integer, Tokens.Byte }, Tokens.Double)));

                    OperatorDeclaredTypesDiv.Add(((new List<string>() { Tokens.Date, Tokens.Currency, Tokens.Double, Tokens.String }, new List<string>() { Tokens.Date, Tokens.Currency, Tokens.Double, Tokens.Single, Tokens.LongLong, Tokens.Long, Tokens.Boolean, Tokens.Integer, Tokens.Byte, Tokens.String }, Tokens.Double)));
                    OperatorDeclaredTypesDiv.Add(((new List<string>() { Tokens.Date, Tokens.Currency, Tokens.Double, Tokens.Single, Tokens.LongLong, Tokens.Long, Tokens.Boolean, Tokens.Integer, Tokens.Byte, Tokens.String }, new List<string>() { Tokens.Date, Tokens.Currency, Tokens.Double, Tokens.String }, Tokens.Double)));
                }

                return RetrieveOperatorDeclaredTypeName(OperatorDeclaredTypesDiv, lhsTypeName, rhsTypeName);
            }

            //Apply '\' Operator exceptions to the default OperatorDeclaredTypesBinary per VBA spec 5.6.9.3
            private string OperatorDeclaredTypeIntegerDivide(string lhsTypeName, string rhsTypeName)
            {
                //Modulo an Integer Division have the same exceptions to the default operatorDeclaredTypes
                return OperatorDeclaredTypeMod(lhsTypeName, rhsTypeName);
            }

            //Apply 'Mod' Operator exceptions to the default OperatorDeclaredTypesBinary per VBA spec 5.6.9.3
            private string OperatorDeclaredTypeMod(string lhsTypeName, string rhsTypeName)
            {
                //Apply 'Mod and \' Operator exceptions to the default OperatorDeclaredTypesBinary per VBA spec 5.6.9.3
                if (OperatorDeclaredTypesMod is null)
                {
                    OperatorDeclaredTypesMod = new List<(List<string> lhsTypes, List<string> rhsTypes, string typeName)>();

                    OperatorDeclaredTypesMod.Add(((new List<string>() { Tokens.Date, Tokens.Currency, Tokens.Double, Tokens.Single, Tokens.String }, new List<string>() { Tokens.Date, Tokens.Currency, Tokens.Double, Tokens.Single, Tokens.LongLong, Tokens.Long, Tokens.Boolean, Tokens.Integer, Tokens.Byte, Tokens.String }, Tokens.Long)));
                    OperatorDeclaredTypesMod.Add(((new List<string>() { Tokens.Date, Tokens.Currency, Tokens.Double, Tokens.Single, Tokens.LongLong, Tokens.Long, Tokens.Boolean, Tokens.Integer, Tokens.Byte, Tokens.String }, new List<string>() { Tokens.Date, Tokens.Currency, Tokens.Double, Tokens.Single, Tokens.String }, Tokens.Long)));
                }

                return RetrieveOperatorDeclaredTypeName(OperatorDeclaredTypesMod, lhsTypeName, rhsTypeName);
            }
        }

        private class OperatorDeclaredTypeRelational : OperatorDeclaredTypeBase, IBinaryOperatorDeclaredType
        {
            public string BinaryOperatorDeclaredType((string lhsTypeName, string rhsTypeName) operands, string opSymbol)
            {
                var operatorDeclaredType = OperatorDeclaredTypeDefaultArithmeticBinary(operands.lhsTypeName, operands.rhsTypeName);
                if (!operands.lhsTypeName.Equals(Tokens.Variant) || !operands.rhsTypeName.Equals(Tokens.Variant))
                {
                    return Tokens.Boolean;
                }
                return Tokens.Variant;
            }
        }

        private class OperatorDeclaredTypeLogical : OperatorDeclaredTypeBase, IBinaryOperatorDeclaredType, IUnaryOperatorDeclaredType
        {
            public string BinaryOperatorDeclaredType((string lhsTypeName, string rhsTypeName) operands, string opSymbol)
            {
                try
                {
                    return OperatorDeclaredTypeDefaultLogical(operands.lhsTypeName, operands.rhsTypeName);
                }
                catch (ArgumentException)
                {
                    return string.Empty;
                }
            }

            public string UnaryOperatorDeclaredType(string operandTypeName, string opSymbol)
            {
                var unaryLong = FloatingPointAndFixedPointNumericTypesStringLongAndDate();
                if (unaryLong.Contains(opSymbol))
                {
                    return Tokens.Long;
                }
                return operandTypeName;
            }
        }
    }
}
