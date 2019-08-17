using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using ExpressiveReflection;
using Moq;
using NLog;

namespace Rubberduck.ComClientLibrary.UnitTesting.Mocks
{
    public class MockArgumentResolver
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// Transform the collection of <see cref="MockArgumentDefinition"/> into a <see cref="IReadOnlyList{T}"/>
        /// </summary>
        /// <remarks>
        /// If a method `Foo` requires one argument, we need to specify the behavior in an expression similar
        /// to this: <c>Mock.Setup(x => x.Foo(It.IsAny())</c>. The class <see cref="It"/> is static so we can
        /// create call expressions directly on it. 
        /// </remarks>
        /// <param name="parameters">Array of <see cref="ParameterInfo"/> returned from the member for which the <see cref="MockArgumentDefinitions"/> applies to</param>
        /// <param name="args">The <see cref="MockArgumentDefinitions"/> collection containing user supplied behavior</param>
        /// <returns>A read-only list containing the <see cref="Expression"/> of arguments</returns>
        public IReadOnlyList<Expression> ResolveParameters(
            IReadOnlyList<ParameterInfo> parameters,
            MockArgumentDefinitions args)
        {
            var argsCount = args?.Count ?? 0;
            if (parameters.Count != argsCount)
            {
                throw new ArgumentOutOfRangeException(nameof(args),
                    $"The method expects {parameters.Count} parameters but only {argsCount} argument definitions were supplied. Setting up the mock's behavior requires that all parameters be filled in.");
            }

            if (parameters.Count == 0)
            {
                return null;
            }

            var resolvedArguments = new List<Expression>();
            for (var i = 0; i < parameters.Count; i++)
            {
                Debug.Assert(args != null, nameof(args) + " != null");

                var parameter = parameters[i];
                var definition = args.Item(i);

                var itType = typeof(It);
                MethodInfo itMemberInfo;

                var parameterType = parameter.ParameterType;
                var itArgumentExpressions = new List<Expression>();
                var typeExpression = Expression.Parameter(parameterType, "x");

                switch (definition.Type)
                {
                    case MockArgumentType.Is:
                        itMemberInfo = itType.GetMethod(nameof(It.Is)).MakeGenericMethod(parameterType);
                        var value = definition.Values[0];
                        if (value != null && value.GetType() != parameterType)
                        {
                            if (TryCast(value, parameterType, out var convertedValue))
                            {
                                value = convertedValue;
                            }
                        }

                        var bodyExpression = Expression.Equal(typeExpression, Expression.Convert(Expression.Constant(value), parameterType));
                        var itLambda = Expression.Lambda(bodyExpression, typeExpression);
                        itArgumentExpressions.Add(itLambda);
                        break;
                    case MockArgumentType.IsAny:
                        itMemberInfo = Reflection.GetMethodExt(itType, nameof(It.IsAny)).MakeGenericMethod(parameterType);
                        break;
                    case MockArgumentType.IsIn:
                        itMemberInfo = Reflection.GetMethodExt(itType, nameof(It.IsIn), typeof(IEnumerable<>)).MakeGenericMethod(parameterType);
                        var arrayInit = Expression.NewArrayInit(parameterType,
                            definition.Values.Select(x => Expression.Convert(Expression.Constant(TryCast(x, parameterType,  out var c) ? c : x), parameterType)));
                        itArgumentExpressions.Add(arrayInit);
                        break;
                    case MockArgumentType.IsInRange:
                        itMemberInfo = Reflection.GetMethodExt(itType, nameof(It.IsInRange), typeof(MethodReflection.T),
                            typeof(MethodReflection.T), typeof(Range)).MakeGenericMethod(parameterType);
                        itArgumentExpressions.Add( Expression.Convert(Expression.Constant(TryCast(definition.Values[0], parameterType, out var from) ? from : definition.Values[0]), parameterType));
                        itArgumentExpressions.Add( Expression.Convert(Expression.Constant(TryCast(definition.Values[1], parameterType, out var to) ? to : definition.Values[1]), parameterType));
                        itArgumentExpressions.Add(definition.Range != null
                            ? Expression.Constant((Range) definition.Range)
                            : Expression.Constant(Range.Inclusive));
                        break;
                    case MockArgumentType.IsNotIn:
                        itMemberInfo = Reflection.GetMethodExt(itType, nameof(It.IsNotIn), typeof(IEnumerable<>)).MakeGenericMethod(parameterType);
                        var notArrayInit = Expression.NewArrayInit(parameterType,
                            definition.Values.Select(x => Expression.Convert(Expression.Constant(TryCast(x, parameterType, out var c) ? c : x), parameterType)));
                        itArgumentExpressions.Add(notArrayInit);
                        break;
                    case MockArgumentType.IsNotNull:
                        itMemberInfo = Reflection.GetMethodExt(itType, nameof(It.IsNotNull)).MakeGenericMethod(parameterType);
                        break;
                    default:
                        throw new ArgumentOutOfRangeException();
                }

                var callExpression = Expression.Call(itMemberInfo, itArgumentExpressions);
                resolvedArguments.Add(callExpression);
            }

            return resolvedArguments;
        }

        private static bool TryCast(object value, Type type, out object convertedValue)
        {
            convertedValue = null;

            try
            {
                convertedValue = Convert.ChangeType(value, type);
            }
            catch
            {
                try
                {
                    convertedValue = VariantConverter.ChangeType(value, type);
                }
                catch
                {
                    Logger.Trace($"Casting failed: the source type was '{value.GetType()}', and the target type wsa '{type.FullName}'");
                }
            }

            return convertedValue != null;
        }
    }
}
