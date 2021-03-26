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

    public interface IRuntimeSetup : IRuntimeMock, IRuntimeCallback, IRuntimeExecute, IRuntimeVerify
    {
        IRuntimeCallback Setup(Expression setupExpression, IReadOnlyDictionary<ParameterExpression, object> forwardedArgs);
        IRuntimeReturns Setup(Expression setupExpression, IReadOnlyDictionary<ParameterExpression, object> forwardedArgs, Type returnType);
    }

    public interface IRuntimeVerify : IRuntimeExecute
    {
        IRuntimeVerify Verify(Expression verifyExpression, ITimes times, IReadOnlyDictionary<ParameterExpression, object> forwardedArgs);
        IRuntimeVerify Verify(Expression verifyExpression, ITimes times, IReadOnlyDictionary<ParameterExpression, object> forwardedArgs, Type returnType);
    }

    public interface IRuntimeCallback : IRuntimeExecute
    {
        IRuntimeSetup Callback(Action callback);
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
        IRuntimeVerify,
        IRuntimeExecute
    {
        private readonly Mock _mock;
        private readonly Type _mockType;
        private readonly ParameterExpression _mockParameterExpression;
        private readonly List<object> _args;
        private readonly List<ParameterExpression> _lambdaArguments;

        private Expression _expression;
        private Type _currentType;

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
            _lambdaArguments = new List<ParameterExpression>
            {
                _mockParameterExpression
            };
        }

        public IRuntimeSetup As(Type targetInterface)
        {
            var asMethodInfo = MockMemberInfos.As(targetInterface);
            _expression = Expression.Call(_mockParameterExpression, asMethodInfo);
            _currentType = asMethodInfo.ReturnType;
            return this;
        }

        public IRuntimeCallback Setup(Expression setupExpression, IReadOnlyDictionary<ParameterExpression, object> forwardedArgs)
        {
            switch (setupExpression.Type.GetGenericArguments().Length)
            {
                case 2:
                    // It's a returning method so we need to use the Func version of Setup and ignore the return.
                    Setup(setupExpression, forwardedArgs, setupExpression.Type.GetGenericArguments()[1]);
                    return this;
                case 1:
                    var setupMethodInfo = MockMemberInfos.Setup(_currentType, null);
                    // Quoting the setup lambda expression ensures that closures will be applied
                    _expression = Expression.Call(_expression, setupMethodInfo, Expression.Quote(setupExpression));
                    _currentType = setupMethodInfo.ReturnType;
                    if (forwardedArgs.Any())
                    {
                        _lambdaArguments.AddRange(forwardedArgs.Keys);
                        _args.AddRange(forwardedArgs.Values);
                    }
                    return this;
                default:
                    throw new NotSupportedException("Setup can only handle 1 or 2 arguments as an input");
            }
        }

        public IRuntimeReturns Setup(Expression setupExpression, IReadOnlyDictionary<ParameterExpression, object> forwardedArgs, Type returnType)
        {
            var setupMethodInfo = MockMemberInfos.Setup(_currentType, returnType);
            // Quoting the setup lambda expression ensures that closures will be applied
            _expression = Expression.Call(_expression, setupMethodInfo, Expression.Quote(setupExpression));
            _currentType = setupMethodInfo.ReturnType;
            if (forwardedArgs.Any())
            {
                _lambdaArguments.AddRange(forwardedArgs.Keys);
                _args.AddRange(forwardedArgs.Values);
            }
            return this;
        }

        public IRuntimeSetup Callback(Action callback)
        {
            var callbackMethodInfo = MockMemberInfos.Callback(_currentType);
            var callbackType = callbackMethodInfo.DeclaringType;

            var valueParameterExpression = Expression.Parameter(callback.GetType(), "value");
            _lambdaArguments.Add(valueParameterExpression);

            var castCallbackExpression = Expression.Convert(_expression, callbackType);
            var callCallbackExpression =
                Expression.Call(castCallbackExpression, callbackMethodInfo, valueParameterExpression);
            _expression = Expression.Lambda(callCallbackExpression, _lambdaArguments);
            _currentType = callbackMethodInfo.ReturnType;
            _args.Add(callback);
            return this;
        }

        public IRuntimeSetup Returns(object value, Type type)
        {
            var returnsMethodInfo = MockMemberInfos.Returns(_currentType);
            var returnsType = returnsMethodInfo.DeclaringType;

            var valueParameterExpression = Expression.Parameter(type, $"value{_args.Count}");
            _lambdaArguments.Add(valueParameterExpression);

            var castReturnExpression = Expression.Convert(_expression, returnsType);
            var returnsCallExpression = Expression.Call(castReturnExpression, returnsMethodInfo, valueParameterExpression);
            _expression = Expression.Lambda(returnsCallExpression, _lambdaArguments);
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

        public IRuntimeVerify Verify(Expression verifyExpression, ITimes times, IReadOnlyDictionary<ParameterExpression, object> forwardedArgs)
        {
            switch (verifyExpression.Type.GetGenericArguments().Length)
            {
                case 2:
                    // It's a returning method so we need to use the Func overload of Setup and ignore the return.
                    Verify(verifyExpression, times, forwardedArgs, verifyExpression.Type.GetGenericArguments()[1]);
                    return this;
                case 1:
                    var verifyMethodInfo = MockMemberInfos.Verify(_currentType);
                    var rdTimes = (Times)times;
                    // Quoting the setup lambda expression ensures that closures will be applied
                    _expression = Expression.Call(_expression, verifyMethodInfo, Expression.Quote(verifyExpression), Expression.Constant(rdTimes.MoqTimes));
                    _currentType = verifyMethodInfo.ReturnType;
                    if (forwardedArgs.Any())
                    {
                        _lambdaArguments.AddRange(forwardedArgs.Keys);
                        _args.AddRange(forwardedArgs.Values);
                    }
                    return this;
                default:
                    throw new NotSupportedException("Verify can only handle 1 or 2 arguments as an input");
            }
        }

        public IRuntimeVerify Verify(Expression verifyExpression, ITimes times, IReadOnlyDictionary<ParameterExpression, object> forwardedArgs, Type returnType)
        {
            var verifyMethodInfo = MockMemberInfos.Verify(_currentType, returnType);
            var rdTimes = (Times)times;
            // Quoting the verify lambda expression ensures that closures will be applied
            _expression = Expression.Call(_expression, verifyMethodInfo, Expression.Quote(verifyExpression), Expression.Constant(rdTimes.MoqTimes));
            _currentType = verifyMethodInfo.ReturnType;
            if (forwardedArgs.Any())
            {
                _lambdaArguments.AddRange(forwardedArgs.Keys);
                _args.AddRange(forwardedArgs.Values);
            }
            return this;
        }
    }
}
