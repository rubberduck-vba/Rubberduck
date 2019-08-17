using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
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

        public static ArgumentSetup[] ArrangeArgumentSetup(MockArgumentType argumentType, Type returnType, object[] value)
        {
            return new[]
            {
                new ArgumentSetup(argumentType, MockArgumentMapper(argumentType), returnType, value)
            };
        }

        public static ArgumentSetup[] ArrangeArgumentSetup(MockArgumentType argumentType, Type returnType, object value)
        {
            return new[]
            {
                new ArgumentSetup(argumentType, MockArgumentMapper(argumentType), returnType, value)
            };
        }
        
        public static AssertData ArrangeAssertData(MethodSelection methodSelection, ArgumentSetup[] argumentSetups)
        {
            var (returnType, methodName) = MethodSelector(methodSelection);

            return new AssertData(
                returnType,
                methodName,
                argumentSetups
            );
        }

        public void AssertMockArgumentResolver(AssertData data)
        {
            var resolver = new MockArgumentResolver();
            var parameterInfos = data.TargetType.GetMethod(data.MethodName)?.GetParameters();

            Assert.IsNotNull(parameterInfos, "Reflection on method failed");

            var mockDefinitions = new MockArgumentDefinitions();
            foreach (var setup in data.ArgumentSetups)
            {
                mockDefinitions.Add(new MockArgumentDefinition(setup.ArgumentType, setup.Value));
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

    public readonly struct ArgumentSetup
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

    public readonly struct AssertData
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
