using NUnit.Framework;
using Rubberduck.Inspections.Concrete.UnreachableCaseInspection;
using Rubberduck.Parsing.Grammar;
using System;

namespace RubberduckTests.Inspections.UnreachableCase
{
    /*
        ParseTreeValue is a support class of the UnreachableCaseInspection

        Test Parameter encoding:
        <operand>?<declaredType>
        If there is no "?<declaredType>", then <operand>'s type is derived by the ParseTreeValue instance.
    */
    [TestFixture]
    public class ParseTreeValueUnitTests
    {
        private const string VALUE_TYPE_SEPARATOR = "?";

        private IUnreachableCaseInspectionFactoryProvider _factoryProvider;
        private IParseTreeValueFactory _valueFactory;
        private IUnreachableCaseInspectionFactoryProvider FactoryProvider
        {
            get
            {
                if (_factoryProvider is null)
                {
                    _factoryProvider = new UnreachableCaseInspectionFactoryProvider();
                }
                return _factoryProvider;
            }
        }

        private IParseTreeValueFactory ValueFactory
        {
            get
            {
                if (_valueFactory is null)
                {
                    _valueFactory = FactoryProvider.CreateIParseTreeValueFactory();
                }
                return _valueFactory;
            }
        }

        [TestCase("2", "2")]
        [TestCase("2.54", "2.54")]
        [TestCase("2.54?Long", "3")]
        [TestCase("2.54?Double", "2.54")]
        [TestCase("2.54?Boolean", "True")]
        [Category("Inspections")]
        public void ParseTreeValue_ConformedToType(string operands, string expectedValueText)
        {
            var value = CreateInspValueFrom(operands);
            Assert.AreEqual(expectedValueText, value.ValueText);
        }

        [Test]
        [Category("Inspections")]
        public void ParseTreeValue_NullInputValue()
        {
            IParseTreeValue test = null;
            try
            {
                test = ValueFactory.Create(null);
                Assert.Fail("Null input to UnreachableCaseInspectionValue did not generate an Argument exception");
            }
            catch (ArgumentException)
            {

            }
            catch
            {
                Assert.Fail("Null input to UnreachableCaseInspectionValue did not generate an exception");
            }
        }

        [TestCase("x", "", "x")]
        [TestCase("x?Variant", "Variant", "x")]
        [TestCase("x?String", "String", "x")]
        [TestCase("x?Double", "Double", "x")]
        [TestCase("x456", "", "x456")]
        [TestCase(@"""x456""", "String", @"""x456""")]
        [TestCase("x456?String", "String", "x456")]
        [TestCase("45E2", "Double", "4500")]
        [TestCase("45E+2", "Double", "4500")]
        [TestCase("45E-2", "Double", "0.45")]
        [TestCase(@"""10.51""", "String", @"""10.51""")]
        [TestCase(@"""What@""", "String", @"""What@""")]
        [TestCase(@"""What!""", "String", @"""What!""")]
        [TestCase(@"""What#""", "String", @"""What#""")]
        [TestCase("What%", "Integer", "What")]
        [TestCase("What&", "Long", "What")]
        [TestCase("What@", "Currency", "What")]
        [TestCase("What!", "Single", "What")]
        [TestCase("What#", "Double", "What")]
        [TestCase("What$", "String", "What")]
        [TestCase("345%", "Integer", "345")]
        [TestCase("45#", "Double", "45")]
        [TestCase("45@", "Currency", "45")]
        [TestCase("45!", "Single", "45")]
        [TestCase("45^", "LongLong", "45")]
        [TestCase("32767", "Integer", "32767")]
        [TestCase("32768", "Long", "32768")]
        [TestCase("2147483647", "Long", "2147483647")]
        [TestCase("2147483648", "Double", "2147483648")]
        [TestCase("&H10", "Integer", "16")]
        [TestCase("&o10", "Integer", "8")]
        [TestCase("&H8000", "Integer", "-32768")]
        [TestCase("&o100000", "Integer", "-32768")]
        [TestCase("&H8000&", "Long", "32768")]
        [TestCase("&o100000&", "Long", "32768")]
        [TestCase("&H10&", "Long", "16")]
        [TestCase("&o10&", "Long", "8")]
        [TestCase("&H80000000", "Long", "-2147483648")]
        [TestCase("&o20000000000", "Long", "-2147483648")]
        [TestCase("&H80000000^", "LongLong", "2147483648")]
        [TestCase("&o20000000000^", "LongLong", "2147483648")]
        [TestCase("&H10^", "LongLong", "16")]
        [TestCase("&o10^", "LongLong", "8")]
        [TestCase("&HFFFFFFFFFFFFFFFF^", "LongLong", "-1")]
        [TestCase("&o1777777777777777777777^", "LongLong", "-1")]
        [Category("Inspections")]
        public void ParseTreeValue_VariableType(string operands, string expectedTypeName, string expectedValueText)
        {
            var value = CreateInspValueFrom(operands);
            Assert.AreEqual(expectedTypeName, value.TypeName);
            Assert.AreEqual(expectedValueText, value.ValueText);
        }

        [TestCase("45.5?Double", "Double", "45.5")]
        [TestCase("45.5?Currency", "Currency", "45.5")]
        [TestCase(@"""45E2""?Long", "Long", "4500")]
        [TestCase(@"""95E-2""?Double", "Double", "0.95")]
        [TestCase(@"""95E-2""?Byte", "Byte", "1")]
        [TestCase("True?Double", "Double", "-1")]
        [TestCase("True?Long", "Long", "-1")]
        [TestCase("&H10", "Integer", "16")]
        [TestCase("&o10", "Integer", "8")]
        [TestCase("&H8000", "Integer", "-32768")]
        [TestCase("&o100000", "Integer", "-32768")]
        [TestCase("&H8000", "Long", "32768")]
        [TestCase("&o100000", "Long", "32768")]
        [TestCase("&H10", "Long", "16")]
        [TestCase("&o10", "Long", "8")]
        [TestCase("&H80000000", "Long", "-2147483648")]
        [TestCase("&o20000000000", "Long", "-2147483648")]
        [TestCase("&H80000000", "LongLong", "2147483648")]
        [TestCase("&o20000000000", "LongLong", "2147483648")]
        [TestCase("&H10", "LongLong", "16")]
        [TestCase("&o10", "LongLong", "8")]
        [TestCase("&HFFFFFFFFFFFFFFFF", "LongLong", "-1")]
        [TestCase("&o1777777777777777777777", "LongLong", "-1")]
        [Category("Inspections")]
        public void ParseTreeValue_ConformToType(string operands, string conformToType, string expectedValueText)
        {
            var value = CreateInspValueFrom(operands, conformToType);

            Assert.AreEqual(conformToType, value.TypeName);
            Assert.AreEqual(expectedValueText, value.ValueText);
        }

        [TestCase("False", "False")]
        [TestCase("True", "True")]
        [TestCase("-1", "True")]
        [TestCase("x < 5.55", "x < 5.55")]
        [Category("Inspections")]
        public void ParseTreeValue_ConvertToBoolText(string input, string expected)
        {
            var result = ValueFactory.Create(input);
            IParseTreeValue coerced = null;
            if (result.ParsesToConstantValue)
            {
                result.TryLetCoerce(Tokens.Boolean, out coerced);
            }
            else
            {
                coerced = result;
            }
            Assert.IsNotNull(coerced, $"Type conversion to {Tokens.Boolean} return null interface");
            Assert.AreEqual(expected, coerced.ValueText);
        }

        [TestCase("Yahoo", "Long")]
        [TestCase("Yahoo", "Double")]
        [TestCase("Yahoo", "Boolean")]
        [Category("Inspections")]
        public void ParseTreeValue_ConvertToType(string input, string convertToTypeName)
        {
            var result = ValueFactory.CreateDeclaredType(input, convertToTypeName);
            Assert.IsNotNull(result, $"Type conversion to {convertToTypeName} return null interface");
            Assert.AreEqual("Yahoo", result.ValueText);
        }

        [TestCase("NaN", "String")]
        [Category("Inspections")]
        public void ParseTreeValue_ConvertToNanType(string input, string convertToTypeName)
        {
            var result = ValueFactory.CreateDeclaredType(input, convertToTypeName);
            Assert.IsNotNull(result, $"Type conversion to {convertToTypeName} return null interface");
            Assert.AreEqual("NaN", result.ValueText);
        }

        [TestCase(@"""W#hat#""", "String", @"""W#hat#""")]
        [TestCase(@"""#W#hat#""", "String", @"""#W#hat#""")]    //Like Date
        [Category("Inspections")]
        public void ParseTreeValue_LikeATypeHint(string operands, string expectedTypeName, string expectedValueText)
        {
            var value = CreateInspValueFrom(operands);
            Assert.AreEqual(expectedTypeName, value.TypeName);
            Assert.AreEqual(expectedValueText, value.ValueText);
        }

        [TestCase("#1/4/2005#", "Date")]
        [TestCase("#4-jan-2006#", "Date")]
        [TestCase("#4-jan#", "Date")]
        [TestCase(@"""#I'mNotADateType0#""", "String")]
        [Category("Inspections")]
        public void ParseTreeValue_DateTypeLiteral(string literal, string expectedTypeName)
        {
            var ptValue = ValueFactory.Create(literal);
            Assert.AreEqual(expectedTypeName, ptValue.TypeName);
        }

        [Test]
        [Category("Inspections")]
        public void ParseTreeValue_DateTypeDeclared()
        {
            var ptValue = ValueFactory.CreateDate("#1/4/2005#");
            Assert.AreEqual(Tokens.Date, ptValue.TypeName);
        }

        private IParseTreeValue CreateInspValueFrom(string valAndType, string conformTo = null)
        {
            var value = valAndType;
            if (valAndType.Contains(VALUE_TYPE_SEPARATOR))
            {
                var args = valAndType.Split(new string[] { VALUE_TYPE_SEPARATOR }, StringSplitOptions.None);
                value = args[0];
                string declaredType = args[1].Equals(string.Empty) ? null : args[1];
                if (conformTo is null)
                {
                    return declaredType is null ? ValueFactory.Create(value) 
                        : ValueFactory.CreateDeclaredType(value, declaredType);
                }
                else
                {
                    return declaredType is null ? ValueFactory.CreateDeclaredType(value, conformTo)
                        : ValueFactory.CreateDeclaredType(value, declaredType);
                }
            }
            return conformTo is null ? ValueFactory.Create(value)
                : ValueFactory.CreateDeclaredType(value, conformTo);
        }
    }
}
