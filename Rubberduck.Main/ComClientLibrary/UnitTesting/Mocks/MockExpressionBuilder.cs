using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Linq.Expressions;
using Moq;

namespace Rubberduck.ComClientLibrary.UnitTesting.Mocks
{
    public interface IRuntimeMock
    {
        IRuntimeSetup As(Type targetInterface);
    }

    public interface IRuntimeSetup : IRuntimeMock, IRuntimeCallback, IRuntimeExecute
    {
        IRuntimeCallback Setup(Expression setupExpression);
        IRuntimeReturns Setup(Expression setupExpression, Type returnType);
    }

    public interface IRuntimeCallback : IRuntimeExecute
    {
        IRuntimeSetup Callback(Action Callback);
    }

    public interface IRuntimeReturns : IRuntimeExecute
    {
        IRuntimeSetup Returns(object value, Type type);
    }

    public interface IRuntimeExecute
    {
        object Execute();
    }

    // Some interfaces are already implemented by others but we list them all explicitly for clarity
    [SuppressMessage("ReSharper", "RedundantExtendsListEntry")]
    public class MockExpressionBuilder :
        IRuntimeMock,
        IRuntimeSetup,
        IRuntimeCallback,
        IRuntimeReturns,
        IRuntimeExecute
    {
        private readonly Mock _mock;
        private readonly Type _mockType;
        private readonly ParameterExpression _mockParameterExpression;
        private Expression _expression;
        private Type _currentType;
        private List<object> _args;

        public static IRuntimeMock Create(Mock runtimeMock)
        {
            return new MockExpressionBuilder(runtimeMock);
        }

        private MockExpressionBuilder(Mock runtimeMock)
        {
            _mock = runtimeMock;
            _mockType = _mock.GetType();
            _mockParameterExpression = Expression.Parameter(_mockType, "mock");
            _args = new List<object>();
        }

        public IRuntimeSetup As(Type targetInterface)
        {
            var asMethodInfo = MockMemberInfos.As(targetInterface);
            _expression = Expression.Call(_mockParameterExpression, asMethodInfo);
            _currentType = asMethodInfo.ReturnType;
            return this;
        }

        public IRuntimeCallback Setup(Expression setupExpression)
        {
            var setupMethodInfo = MockMemberInfos.Setup(_currentType, null);
            _expression = Expression.Call(_expression, setupMethodInfo, setupExpression);
            _currentType = setupMethodInfo.ReturnType;
            return this;
        }

        public IRuntimeReturns Setup(Expression setupExpression, Type returnType)
        {
            var setupMethodInfo = MockMemberInfos.Setup(_currentType, returnType);
            _expression = Expression.Call(_expression, setupMethodInfo, setupExpression);
            _currentType = setupMethodInfo.ReturnType;
            return this;
        }

        public IRuntimeSetup Callback(Action callback)
        {
            var callbackMethodInfo = MockMemberInfos.Callback(_currentType);
            var callbackType = callbackMethodInfo.DeclaringType;

            var valueParameterExpression = Expression.Parameter(callback.GetType(), "value");

            var castCallbackExpression = Expression.Convert(_expression, callbackType);
            var callCallbackExpression =
                Expression.Call(castCallbackExpression, callbackMethodInfo, valueParameterExpression);
            _expression = Expression.Lambda(callCallbackExpression, _mockParameterExpression,
                valueParameterExpression);
            _currentType = callbackMethodInfo.ReturnType;
            _args.Add(callback);
            return this;
        }

        public IRuntimeSetup Returns(object value, Type type)
        {
            var returnsMethodInfo = MockMemberInfos.Returns(_currentType);
            var returnsType = returnsMethodInfo.DeclaringType;

            var valueParameterExpression = Expression.Parameter(type, $"value{_args.Count}");

            var castReturnExpression = Expression.Convert(_expression, returnsType);
            var returnsCallExpression = Expression.Call(castReturnExpression, returnsMethodInfo, valueParameterExpression);
            _expression = Expression.Lambda(returnsCallExpression, _mockParameterExpression, valueParameterExpression);
            _currentType = returnsMethodInfo.ReturnType;
            _args.Add(value);
            return this;
        }

        public object Execute()
        {
            var args = _args.Count >= 0 ? new List<object> {_mock}.Concat(_args).ToArray() : _args.ToArray();

            return _expression.NodeType == ExpressionType.Lambda 
                ? ((LambdaExpression)_expression).Compile().DynamicInvoke(args)
                : Expression.Lambda(_expression, _mockParameterExpression).Compile().DynamicInvoke(args);
        }
    }
}
