using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using Moq;
using NUnit.Framework;
using Rubberduck.ComClientLibrary.UnitTesting.Mocks;

namespace RubberduckTests.ComMock
{
    [TestFixture]
    [Category("ComMocks.MockExpressionBuilderTests")]
    public class MockExpressionBuilderTests
    {
        [Test]
        public void As_Compiles()
        {
            var mock = new Mock<ITest1>();
            var builder = MockExpressionBuilder.Create(mock);

            builder.As(typeof(ITest1))
                   .Execute();
        }

        [Test]
        public void Setup_Void_Method_Compiles()
        {
            var mock = new Mock<ITest1>();
            var builder = MockExpressionBuilder.Create(mock);
            var expression = ArrangeSetupDoExpression();

            builder.As(typeof(ITest1))
                .Setup(expression, ArrangeForwardedArgs())
                .Execute();
        }

        [Test]
        public void Setup_ReturningMethod_WithReturnIgnored_Compiles()
        {
            var mock = new Mock<ITest1>();
            var builder = MockExpressionBuilder.Create(mock);
            var expression = ArrangeSetupDoThisExpression();

            builder.As(typeof(ITest1))
                .Setup(expression, ArrangeForwardedArgs())
                .Execute();
        }

        [Test]
        public void Setup_ReturningMethod_Compiles()
        {
            var mock = new Mock<ITest1>();
            var builder = MockExpressionBuilder.Create(mock);
            var expression = ArrangeSetupDoThisExpression();

            builder.As(typeof(ITest1))
                   .Setup(expression, ArrangeForwardedArgs(), typeof(int))
                   .Execute();
        }

        [Test]
        public void SetupWithReturns_Compiles()
        {
            const int expected = 42;
            var mock = new Mock<ITest1>();
            var builder = MockExpressionBuilder.Create(mock);
            var expression = ArrangeSetupDoThisExpression();

            builder.As(typeof(ITest1))
                .Setup(expression, ArrangeForwardedArgs(), typeof(int))
                .Returns(expected, typeof(int))
                .Execute();

            Assert.AreEqual(expected, mock.Object.DoThis());
        }

        [Test]
        public void SetupWithCallback_Compiles()
        {
            const int expected = 42;
            var actual = 0;
            var action = new Action(() => { actual = expected; });
            var mock = new Mock<ITest1>();
            var builder = MockExpressionBuilder.Create(mock);
            var expression = ArrangeSetupDoExpression();

            builder.As(typeof(ITest1))
                .Setup(expression, ArrangeForwardedArgs())
                .Callback(action)
                .Execute();

            mock.Object.Do();

            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void Verify_Compiles()
        {
            const int expected = 1;
            var expectedTimes = Moq.Times.Once().ToRubberduckTimes();

            var mock = new Mock<ITest1>();
            var builder = MockExpressionBuilder.Create(mock);
            var expression = ArrangeSetupDoExpression();

            builder.As(typeof(ITest1))
                .Verify(expression, expectedTimes, ArrangeForwardedArgs())
                .Execute();

            mock.Object.Do();
            Assert.AreEqual(expected, mock.Invocations.Count);
        }

        private static IReadOnlyDictionary<ParameterExpression, object> ArrangeForwardedArgs()
        {
            return new Dictionary<ParameterExpression, object>();
        }

        private static Expression ArrangeSetupDoExpression()
        {
            // x => x.Do()

            var typeParameterExpression = Expression.Parameter(typeof(ITest1), "x");
            var methodInfo = typeof(ITest1).GetMethod(nameof(ITest1.Do));
            var callExpression = Expression.Call(typeParameterExpression, methodInfo);
            return Expression.Lambda(callExpression, typeParameterExpression);
        }

        private static Expression ArrangeSetupDoThisExpression()
        {
            // x => x.DoThis()

            var typeParameterExpression = Expression.Parameter(typeof(ITest1), "x");
            var methodInfo = typeof(ITest1).GetMethod(nameof(ITest1.DoThis));
            var callExpression = Expression.Call(typeParameterExpression, methodInfo);
            return Expression.Lambda(callExpression, typeParameterExpression);
        }
    }
}
