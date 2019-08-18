using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Moq;
using NUnit.Framework;
using Rubberduck.ComClientLibrary.UnitTesting.Mocks;

namespace RubberduckTests.ComMock
{
    [TestFixture]
    [Category("ComMocks.MockArgumentResolverTests")]
    public class MockArgumentResolverTests
    {
        [Test]
        public void Resolve_Args_Null_Returns_Null()
        {
            var resolver = ArrangeMockArgumentResolver();
            var results = resolver.ResolveArgs(null);

            Assert.IsNull(results);
        }

        [Test]
        public void Resolve_Args_Missing_Returns_Null()
        {
            var resolver = ArrangeMockArgumentResolver();
            var arg = Missing.Value;
            var results = resolver.ResolveArgs(arg);

            Assert.IsNull(results);
        }

        [Test]
        public void Resolve_Args_Missing_In_Array_Returns_Null()
        {
            var resolver = ArrangeMockArgumentResolver();
            object[] arg = {Missing.Value};
            var results = resolver.ResolveArgs(arg);

            Assert.IsNull(results);
        }

        [Test]
        public void Resolve_Args_Two_Missing_Returns_Two_IsAny()
        {
            var resolver = ArrangeMockArgumentResolver();
            var arg = new[] {Missing.Value, Missing.Value};
            var results = resolver.ResolveArgs(arg);

            Assert.AreEqual(2, results.Count);
            foreach (var definition in results)
            {
                Assert.AreEqual(MockArgumentType.IsAny, definition.Type);    
            }
        }

        [Test]
        public void Resolve_Args_Two_Nulls_Returns_Two_IsAny()
        {
            var resolver = ArrangeMockArgumentResolver();
            var arg = new object[] {null, null};
            var results = resolver.ResolveArgs(arg);

            Assert.AreEqual(2, results.Count);
            foreach (var definition in results)
            {
                Assert.AreEqual(MockArgumentType.IsAny, definition.Type);
            }
        }

        [Test]
        public void Resolve_Args_Definition_Returns_Definitions()
        {
            var resolver = ArrangeMockArgumentResolver();
            var arg = MockArgumentDefinition.CreateIs(1);
            var results = resolver.ResolveArgs(arg);

            Assert.AreEqual(1, results.Count);
            foreach (var definition in results)
            {
                Assert.AreEqual(MockArgumentType.Is, definition.Type);
                Assert.AreEqual(1, definition.Values.Single());
            }
        }

        [Test]
        public void Resolve_Args_Definition_In_Array_Returns_Definitions()
        {
            var resolver = ArrangeMockArgumentResolver();
            var arg = new[] { MockArgumentDefinition.CreateIs(1) };
            var results = resolver.ResolveArgs(arg);

            Assert.AreEqual(1, results.Count);
            foreach (var definition in results)
            {
                Assert.AreEqual(MockArgumentType.Is, definition.Type);
                Assert.AreEqual(1, definition.Values.Single());
            }
        }

        [Test]
        public void Resolve_Args_Two_Definition_Returns_Definitions()
        {
            var resolver = ArrangeMockArgumentResolver();
            var arg = new[] { MockArgumentDefinition.CreateIs(1), MockArgumentDefinition.CreateIs(2) };
            var results = resolver.ResolveArgs(arg);

            Assert.AreEqual(2, results.Count);

            var i = 1;
            foreach (var definition in results)
            {
                Assert.AreEqual(MockArgumentType.Is, definition.Type);
                Assert.AreEqual(i++, definition.Values.Single());
            }
        }

        [Test]
        public void Resolve_Args_Definitions_Returns_Itself()
        {
            var definitions = new MockArgumentDefinitions
            {
                MockArgumentDefinition.CreateIs(1),
                MockArgumentDefinition.CreateIs(2)
            };

            var resolver = ArrangeMockArgumentResolver();
            var results = resolver.ResolveArgs(definitions);

            Assert.AreSame(definitions, results);
        }

        [Test]
        public void Resolve_Args_Objects_Returns_Definitions()
        {
            var resolver = ArrangeMockArgumentResolver();
            var arg = new object[] {1, 2}; // must be boxed since we take them as variants from COM
            var results = resolver.ResolveArgs(arg);

            Assert.AreEqual(2, results.Count);
            foreach (var definition in results)
            {
                Assert.AreEqual(MockArgumentType.Is, definition.Type);
            }
        }

        [Test]
        [TestCase(1)]
        [TestCase("1")]
        [TestCase("")]
        [TestCase(1.0)]
        public void Resolve_Args_Single_Argument_Returns_Definitions(object arg)
        {
            var resolver = ArrangeMockArgumentResolver();
            var results = resolver.ResolveArgs(arg);

            Assert.AreEqual(1, results.Count);
            foreach (var definition in results)
            {
                Assert.AreEqual(MockArgumentType.Is, definition.Type);
            }
        }

        [Test]
        public void Resolve_Args_Single_Array_Returns_In_Definition()
        {
            var array = new object[] {1, 3, 5}; // must be boxed because we get it as variant from COM
            var resolver = ArrangeMockArgumentResolver();
            var args = new object[] {array}; // arrays must be double-wrapped
            var results = resolver.ResolveArgs(args);

            Assert.AreEqual(1, results.Count);
            var result = results.Single();
            Assert.AreEqual(MockArgumentType.IsIn, result.Type);
            Assert.AreEqual(array, result.Values);
        }

        [Test]
        public void Resolve_Args_Two_Array_Returns_In_Definition()
        {
            var array1 = new object[] {1, 3, 5}; // must be boxed because we get it as variant from COM
            var array2 = new object[] {2, 4, 6};
            var resolver = ArrangeMockArgumentResolver();
            var args = new object[] { array1, array2 }; // arrays must be double-wrapped
            var results = resolver.ResolveArgs(args);

            Assert.AreEqual(2, results.Count);
            var i = 0;
            foreach (var definition in results)
            {
                Assert.AreEqual(MockArgumentType.IsIn, definition.Type);
                Assert.AreEqual(args[i++], definition.Values);
            }
        }

        [Test]
        public void Resolve_Args_Mixed_Array_And_Single_Returns_In_Definition()
        {
            var array = new object[] { 1, 3, 5 }; // must be boxed because we get it as variant from COM
            object singleObject = 2;
            var resolver = ArrangeMockArgumentResolver();
            var args = new object[] {array, singleObject}; // arrays must be double-wrapped
            var results = resolver.ResolveArgs(args);

            Assert.AreEqual(2, results.Count);
            Assert.AreEqual(MockArgumentType.IsIn, results.First().Type);
            Assert.AreEqual(array, results.First().Values);
            Assert.AreEqual(MockArgumentType.Is, results.Last().Type);
            Assert.AreEqual(singleObject, results.Last().Values.Single());
        }

        [Test]
        public void Resolve_Args_Mixed_Single_And_Array_Returns_In_Definition()
        {
            var array = new object[] { 1, 3, 5 }; // must be boxed because we get it as variant from COM
            object singleObject = 2;
            var resolver = ArrangeMockArgumentResolver();
            var args = new object[] {singleObject, array}; // arrays must be double-wrapped
            var results = resolver.ResolveArgs(args);

            Assert.AreEqual(2, results.Count);
            Assert.AreEqual(MockArgumentType.Is, results.First().Type);
            Assert.AreEqual(singleObject, results.First().Values.Single());
            Assert.AreEqual(MockArgumentType.IsIn, results.Last().Type);
            Assert.AreEqual(array, results.Last().Values);
        }

        [Test]
        [TestCase(1, 1)]
        [TestCase(1, "1")]
        [TestCase(1, "")]
        [TestCase(1, 1.0)]
        [TestCase("1", 1)]
        [TestCase("1", "1")]
        [TestCase("1", "")]
        [TestCase("1", 1.0)]
        [TestCase("", 1)]
        [TestCase("", "1")]
        [TestCase("", "")]
        [TestCase("", 1.0)]
        [TestCase(1.0, 1)]
        [TestCase(1.0, "1")]
        [TestCase(1.0, "")]
        [TestCase(1.0, 1.0)]
        public void Resolve_Args_Two_Argument_Returns_Definitions(object arg1, object arg2)
        {
            var resolver = ArrangeMockArgumentResolver();
            var args = new[] {arg1, arg2};
            var results = resolver.ResolveArgs(args);

            Assert.AreEqual(2, results.Count);

            var i = 0;
            foreach (var definition in results)
            {
                var arg = args.ElementAt(i++);
                Assert.AreEqual(MockArgumentType.Is, definition.Type);
                Assert.AreEqual(arg,definition.Values.Single());
            }
        }

        [Test]

        [TestCase(MethodSelection.DoInt, MockArgumentType.IsAny, typeof(int), 1)]
        [TestCase(MethodSelection.DoInt, MockArgumentType.IsAny, typeof(int), 2.2)]
        [TestCase(MethodSelection.DoInt, MockArgumentType.IsAny, typeof(int), "1")]
        [TestCase(MethodSelection.DoInt, MockArgumentType.IsAny, typeof(int), null)]
        [TestCase(MethodSelection.DoString, MockArgumentType.IsAny, typeof(string), 1)]
        [TestCase(MethodSelection.DoString, MockArgumentType.IsAny, typeof(string), 2.2)]
        [TestCase(MethodSelection.DoString, MockArgumentType.IsAny, typeof(string), "1")]
        [TestCase(MethodSelection.DoString, MockArgumentType.IsAny, typeof(string), null)]
        [TestCase(MethodSelection.DoObject, MockArgumentType.IsAny, typeof(object), 1)]
        [TestCase(MethodSelection.DoObject, MockArgumentType.IsAny, typeof(object), 2.2)]
        [TestCase(MethodSelection.DoObject, MockArgumentType.IsAny, typeof(object), "1")]
        [TestCase(MethodSelection.DoObject, MockArgumentType.IsAny, typeof(object), null)]

        [TestCase(MethodSelection.DoInt, MockArgumentType.Is, typeof(int), 1)]
        [TestCase(MethodSelection.DoInt, MockArgumentType.Is, typeof(int), 2.2)]
        [TestCase(MethodSelection.DoInt, MockArgumentType.Is, typeof(int), "1")]
        [TestCase(MethodSelection.DoInt, MockArgumentType.Is, typeof(int), null)]
        [TestCase(MethodSelection.DoString, MockArgumentType.Is, typeof(string), 1)]
        [TestCase(MethodSelection.DoString, MockArgumentType.Is, typeof(string), 2.2)]
        [TestCase(MethodSelection.DoString, MockArgumentType.Is, typeof(string), "1")]
        [TestCase(MethodSelection.DoString, MockArgumentType.Is, typeof(string), null)]
        [TestCase(MethodSelection.DoObject, MockArgumentType.Is, typeof(object), 1)]
        [TestCase(MethodSelection.DoObject, MockArgumentType.Is, typeof(object), 2.2)]
        [TestCase(MethodSelection.DoObject, MockArgumentType.Is, typeof(object), "1")]
        [TestCase(MethodSelection.DoObject, MockArgumentType.Is, typeof(object), null)]

        [TestCase(MethodSelection.DoInt, MockArgumentType.IsNotNull, typeof(int), 1)]
        [TestCase(MethodSelection.DoInt, MockArgumentType.IsNotNull, typeof(int), 2.2)]
        [TestCase(MethodSelection.DoInt, MockArgumentType.IsNotNull, typeof(int), "1")]
        [TestCase(MethodSelection.DoInt, MockArgumentType.IsNotNull, typeof(int), null)]
        [TestCase(MethodSelection.DoString, MockArgumentType.IsNotNull, typeof(string), 1)]
        [TestCase(MethodSelection.DoString, MockArgumentType.IsNotNull, typeof(string), 2.2)]
        [TestCase(MethodSelection.DoString, MockArgumentType.IsNotNull, typeof(string), "1")]
        [TestCase(MethodSelection.DoString, MockArgumentType.IsNotNull, typeof(string), null)]
        [TestCase(MethodSelection.DoObject, MockArgumentType.IsNotNull, typeof(object), 1)]
        [TestCase(MethodSelection.DoObject, MockArgumentType.IsNotNull, typeof(object), 2.2)]
        [TestCase(MethodSelection.DoObject, MockArgumentType.IsNotNull, typeof(object), "1")]
        [TestCase(MethodSelection.DoObject, MockArgumentType.IsNotNull, typeof(object), null)]

        [TestCase(MethodSelection.DoInt, MockArgumentType.IsIn, typeof(int), new[] {1, 3, 5})]
        [TestCase(MethodSelection.DoInt, MockArgumentType.IsIn, typeof(int), new[] {2.2, 4.4, 6.6})]
        [TestCase(MethodSelection.DoInt, MockArgumentType.IsIn, typeof(int), new[] {"1", "3", "5"})]
        [TestCase(MethodSelection.DoString, MockArgumentType.IsIn, typeof(string), new[] { 1, 3, 5 })]
        [TestCase(MethodSelection.DoString, MockArgumentType.IsIn, typeof(string), new[] { 2.2, 4.4, 6.6 })]
        [TestCase(MethodSelection.DoString, MockArgumentType.IsIn, typeof(string), new[] { "1", "3", "5" })]
        [TestCase(MethodSelection.DoObject, MockArgumentType.IsIn, typeof(object), new[] { 1, 3, 5 })]
        [TestCase(MethodSelection.DoObject, MockArgumentType.IsIn, typeof(object), new[] { 2.2, 4.4, 6.6 })]
        [TestCase(MethodSelection.DoObject, MockArgumentType.IsIn, typeof(object), new[] { "1", "3", "5" })]

        [TestCase(MethodSelection.DoInt, MockArgumentType.IsNotIn, typeof(int), new[] { 1, 3, 5 })]
        [TestCase(MethodSelection.DoInt, MockArgumentType.IsNotIn, typeof(int), new[] { 2.2, 4.4, 6.6 })]
        [TestCase(MethodSelection.DoInt, MockArgumentType.IsNotIn, typeof(int), new[] { "1", "3", "5" })]
        [TestCase(MethodSelection.DoString, MockArgumentType.IsNotIn, typeof(string), new[] { 1, 3, 5 })]
        [TestCase(MethodSelection.DoString, MockArgumentType.IsNotIn, typeof(string), new[] { 2.2, 4.4, 6.6 })]
        [TestCase(MethodSelection.DoString, MockArgumentType.IsNotIn, typeof(string), new[] { "1", "3", "5" })]
        [TestCase(MethodSelection.DoObject, MockArgumentType.IsNotIn, typeof(object), new[] { 1, 3, 5 })]
        [TestCase(MethodSelection.DoObject, MockArgumentType.IsNotIn, typeof(object), new[] { 2.2, 4.4, 6.6 })]
        [TestCase(MethodSelection.DoObject, MockArgumentType.IsNotIn, typeof(object), new[] { "1", "3", "5" })]

        // Cannot use objects for IsInRange because it does not have IComparable
        [TestCase(MethodSelection.DoInt, MockArgumentType.IsInRange, typeof(int), new[] { 1, 5 })]
        [TestCase(MethodSelection.DoInt, MockArgumentType.IsInRange, typeof(int), new[] { 2.2, 6.6 })]
        [TestCase(MethodSelection.DoInt, MockArgumentType.IsInRange, typeof(int), new[] { "1", "5" })]
        [TestCase(MethodSelection.DoString, MockArgumentType.IsInRange, typeof(string), new[] { 1, 5 })]
        [TestCase(MethodSelection.DoString, MockArgumentType.IsInRange, typeof(string), new[] { 2.2, 6.6 })]
        [TestCase(MethodSelection.DoString, MockArgumentType.IsInRange, typeof(string), new[] { "1", "5" })]

        public void It_SingleParameter_Tests(MethodSelection methodSelection, MockArgumentType argumentType, Type returnType, object value)
        {
            ArgumentSetup[] argumentSetups;
            if (value != null && value.GetType().IsArray)
            {
                var values = ((IEnumerable) value).Cast<object>().ToArray();
                argumentSetups = ArrangeArgumentSetup(argumentType, returnType, values);
            }
            else
            {
                argumentSetups = ArrangeArgumentSetup(argumentType, returnType, value);
            }

            var assertData = ArrangeAssertData(methodSelection, argumentSetups);

            AssertMockArgumentResolver(assertData);
        }

        public static string MockArgumentMapper(MockArgumentType argumentType)
        {
            switch (argumentType)
            {
                case MockArgumentType.Is:
                    return nameof(It.Is);
                case MockArgumentType.IsAny:
                    return nameof(It.IsAny);
                case MockArgumentType.IsIn:
                    return nameof(It.IsIn);
                case MockArgumentType.IsInRange:
                    return nameof(It.IsInRange);
                case MockArgumentType.IsNotIn:
                    return nameof(It.IsNotIn);
                case MockArgumentType.IsNotNull:
                    return nameof(It.IsNotNull);
                default:
                    throw new ArgumentOutOfRangeException(nameof(argumentType), argumentType, null);
            }
        }

        public enum MethodSelection
        {
            DoInt,
            DoString,
            DoObject
        }

        public static (Type type, string name) MethodSelector(MethodSelection selection)
        {
            switch (selection)
            {
                case MethodSelection.DoInt:
                    return (typeof(ITest3), nameof(ITest3.DoInt));
                case MethodSelection.DoString:
                    return (typeof(ITest3), nameof(ITest3.DoString));
                case MethodSelection.DoObject:
                    return (typeof(ITest3), nameof(ITest3.DoObject));
                default:
                    throw new ArgumentException($"Invalid enumeration for {nameof(MethodSelection)}");
            }
        }

        internal static ArgumentSetup[] ArrangeArgumentSetup(MockArgumentType argumentType, Type returnType, object[] value)
        {
            return new[]
            {
                new ArgumentSetup(argumentType, MockArgumentMapper(argumentType), returnType, value)
            };
        }

        internal static ArgumentSetup[] ArrangeArgumentSetup(MockArgumentType argumentType, Type returnType, object value)
        {
            return new[]
            {
                new ArgumentSetup(argumentType, MockArgumentMapper(argumentType), returnType, value)
            };
        }

        internal static AssertData ArrangeAssertData(MethodSelection methodSelection, ArgumentSetup[] argumentSetups)
        {
            var (returnType, methodName) = MethodSelector(methodSelection);

            return new AssertData(
                returnType,
                methodName,
                argumentSetups
            );
        }

        internal static MockArgumentResolver ArrangeMockArgumentResolver()
        {
            return new MockArgumentResolver();
        }

        internal void AssertMockArgumentResolver(AssertData data)
        {
            var resolver = ArrangeMockArgumentResolver();
            var parameterInfos = data.TargetType.GetMethod(data.MethodName)?.GetParameters();

            Assert.IsNotNull(parameterInfos, "Reflection on method failed");

            var mockDefinitions = new MockArgumentDefinitions();
            foreach (var setup in data.ArgumentSetups)
            {
                MockArgumentDefinition definition;
                switch (setup.ArgumentType)
                {
                    case MockArgumentType.Is:
                        definition = MockArgumentDefinition.CreateIs(setup.Value.Single());
                        break;
                    case MockArgumentType.IsAny:
                        definition = MockArgumentDefinition.CreateIsAny();
                        break;
                    case MockArgumentType.IsIn:
                        definition = MockArgumentDefinition.CreateIsIn(setup.Value);
                        break;
                    case MockArgumentType.IsInRange:
                        Assert.AreEqual(2, setup.Value.Length);
                        definition=MockArgumentDefinition.CreateIsInRange(setup.Value[0], setup.Value[1], MockArgumentRange.Inclusive);
                        break;
                    case MockArgumentType.IsNotIn:
                        definition = MockArgumentDefinition.CreateIsNotIn(setup.Value);
                        break;
                    case MockArgumentType.IsNotNull:
                        definition = MockArgumentDefinition.CreateIsNotNull();
                        break;
                    default:
                        throw new ArgumentOutOfRangeException();
                }
                mockDefinitions.Add(definition);
            }

            var expressions = resolver.ResolveParameters(parameterInfos, mockDefinitions);

            Assert.AreEqual(parameterInfos.Length, expressions.Count);

            var i = 0;
            foreach (var expression in expressions)
            {
                var assertData = data.ArgumentSetups.ElementAt(i++);

                Assert.AreEqual(assertData.ReturnType, expression.Type);
                Assert.IsTrue(expression.ToString().StartsWith(string.Concat(assertData.ItType, "(")));
            }
        }
    }

    internal readonly struct ArgumentSetup
    {
        public MockArgumentType ArgumentType { get; }
        public string ItType { get; }
        public Type ReturnType { get; }
        public object[] Value { get; }
        
        public ArgumentSetup(MockArgumentType argumentType, string itType, Type returnType, object[] value)
        {
            ArgumentType = argumentType;
            ItType = itType;
            ReturnType = returnType;
            Value = value;
        }

        public ArgumentSetup(MockArgumentType argumentType, string itType, Type returnType, object value)
        {
            ArgumentType = argumentType;
            ItType = itType;
            ReturnType = returnType;
            {
                Value = new[] { value };
            }
        }
    }

    internal readonly struct AssertData
    {
        public Type TargetType { get; }
        public string MethodName { get; }
        public IEnumerable<ArgumentSetup> ArgumentSetups { get; }
        public AssertData(Type targetType, string methodName, IEnumerable<ArgumentSetup> argumentSetups)
        {
            TargetType = targetType;
            MethodName = methodName;
            ArgumentSetups = argumentSetups;
        }
    }
}
