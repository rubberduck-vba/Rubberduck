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
    internal class SetupArgumentResolver
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// Converts a variant args into the <see cref="SetupArgumentDefinitions"/> collection. This supports calls from COM
        /// using the Variant data type.
        /// </summary>
        /// <remarks>
        /// The procedure needs to handle the following cases where the variant...:
        ///   1) contains a single value
        ///   2) contains an Array() of values
        ///   3) wraps a single <see cref="SetupArgumentDefinition"/>
        ///   4) points to a <see cref="SetupArgumentDefinitions"/> collection.
        ///   5) wraps a single <see cref="Missing"/> object in which case we return a null
        ///   6) wraps an array of single <see cref="Missing"/> object in which case we return a null
        ///
        /// We must ensure that the arrays are resolved before calling the
        /// single object wrapper to ensure we don't end up wrapping the 
        /// arrays as a single value; do not change the switch order willy-nilly.
        ///
        /// We also need to handle the special cases with <see cref="Missing"/>, because
        /// the methods <see cref="IComMock.SetupWithReturns"/> and <see cref="IComMock.SetupWithCallback"/>
        /// will marshal the Args parameter as a variant, which means we receive it as <see cref="Missing"/>,
        /// not as <c>null</c> if it is omitted. 
        /// </remarks>
        /// <param name="args">Should be a COM Variant that can be cast into valid values as explained in the remarks</param>
        /// <returns>A <see cref="SetupArgumentDefinitions"/> collection or null</returns>
        public SetupArgumentDefinitions ResolveArgs(object args)
        {
            switch (args)
            {
                case Missing missing:
                    return null;
                case SetupArgumentDefinitions definitions:
                    return definitions;
                case SetupArgumentDefinition definition:
                    return WrapArgumentDefinitions(definition);
                case object[] objects:
                    if (objects.Length == 1 && objects[0] is Missing)
                    {
                        return null;
                    }
                    return WrapArgumentDefinitions(objects);
                case object singleObject:
                    return WrapArgumentDefinitions(singleObject);
                default:
                    return null;
            }
        }

        private static SetupArgumentDefinitions WrapArgumentDefinitions(object singleObject)
        {
            var list = new SetupArgumentDefinitions();
            var isDefinition = SetupArgumentDefinition.CreateIs(singleObject);
            list.Add(isDefinition);
            return list;
        }

        private static SetupArgumentDefinitions WrapArgumentDefinitions(object[] objects)
        {
            var list = new SetupArgumentDefinitions();
            foreach (var item in objects)
            {
                switch (item)
                {
                    case SetupArgumentDefinition argumentDefinition:
                        list.Add(argumentDefinition);
                        break;
                    case object[] arrayObjects:
                        var inDefinition = SetupArgumentDefinition.CreateIsIn(arrayObjects);
                        list.Add(inDefinition);
                        break;
                    case Missing missing:
                        list.Add(SetupArgumentDefinition.CreateIsAny());
                        break;
                    case object singleObject:
                        var isDefinition =
                            SetupArgumentDefinition.CreateIs(singleObject);
                        list.Add(isDefinition);
                        break;
                    case null:
                        list.Add(SetupArgumentDefinition.CreateIsAny());
                        break;
                }
            }

            return list;
        }

        private static SetupArgumentDefinitions WrapArgumentDefinitions(SetupArgumentDefinition setupArgumentDefinition)
        {
            return new SetupArgumentDefinitions
            {
                setupArgumentDefinition
            };
        }

        /// <summary>
        /// Transform the collection of <see cref="SetupArgumentDefinition"/> into a <see cref="IReadOnlyList{T}"/>
        /// </summary>
        /// <remarks>
        /// If a method `Foo` requires one argument, we need to specify the behavior in an expression similar
        /// to this: <c>Mock.Setup(x => x.Foo(It.IsAny())</c>. The class <see cref="It"/> is static so we can
        /// create call expressions directly on it. 
        /// </remarks>
        /// <param name="parameters">Array of <see cref="ParameterInfo"/> returned from the member for which the <see cref="SetupArgumentDefinitions"/> applies to</param>
        /// <param name="args">The <see cref="SetupArgumentDefinitions"/> collection containing user supplied behavior</param>
        /// <returns>A read-only list containing the <see cref="Expression"/> of arguments</returns>
        public (IReadOnlyList<Expression> expressions, IReadOnlyDictionary<ParameterExpression, object> forwardedArgs) ResolveParameters(
            IReadOnlyList<ParameterInfo> parameters,
            SetupArgumentDefinitions args)
        {
            var argsCount = args?.Count ?? 0;
            if (parameters.Count != argsCount)
            {
                throw new ArgumentOutOfRangeException(nameof(args),
                    $"The method expects {parameters.Count} parameters but only {argsCount} argument definitions were supplied. Setting up the mock's behavior requires that all parameters be filled in.");
            }

            if (parameters.Count == 0)
            {
                return (null, null);
            }

            var resolvedArguments = new List<Expression>();
            var forwardedArgs = new Dictionary<ParameterExpression, object>();
            for (var i = 0; i < parameters.Count; i++)
            {
                Debug.Assert(args != null, nameof(args) + " != null");

                var parameter = parameters[i];
                var definition = args.Item(i);

                var (elementType, isRef, isOut) = GetParameterType(parameter);
                var parameterType = parameter.ParameterType;

                Expression setupExpression;
                if (isRef || isOut)
                {
                    setupExpression = BuildPassByRefArgumentExpression(i, definition, parameterType, elementType, ref forwardedArgs);
                }
                else
                {
                    setupExpression = BuildPassByValueArgumentExpression(i, definition, parameterType);
                }

                resolvedArguments.Add(setupExpression);
            }

            return (resolvedArguments, forwardedArgs);
        }

        private Expression BuildPassByValueArgumentExpression(int index, SetupArgumentDefinition definition, Type parameterType)
        {
            var itType = typeof(It);
            MethodInfo itMemberInfo;

            var itArgumentExpressions = new List<Expression>();
            var typeExpression = Expression.Parameter(parameterType, $"p{index:00}");

            switch (definition.Type)
            {
                case SetupArgumentType.Is:
                    itMemberInfo = itType.GetMethods().Single(x => x.Name == nameof(It.Is) && x.IsGenericMethodDefinition && x.GetParameters().All(y => y.ParameterType.GetGenericArguments().All(z => z.GetGenericTypeDefinition() == typeof(Func<,>)))).MakeGenericMethod(parameterType);
                    var value = definition.Values[0];
                    if (value != null && value.GetType() != parameterType)
                    {
                        if (TryCast(value, parameterType, out var convertedValue))
                        {
                            value = convertedValue;
                        }
                    }

                    Expression bodyExpression;
                    if (parameterType == typeof(object))
                    {
                        // Avoid incorrectly comparing by reference
                        var equalsInfo = Reflection.GetMethod(() => default(object).Equals(default(object)));

                        bodyExpression = Expression.Call(typeExpression, equalsInfo,
                            Expression.Convert(Expression.Constant(value), parameterType));
                    }
                    else
                    {
                        bodyExpression = Expression.Equal(typeExpression, Expression.Convert(Expression.Constant(value), parameterType));
                    }
                    var itLambda = Expression.Lambda(bodyExpression, typeExpression);
                    itArgumentExpressions.Add(Expression.Quote(itLambda));
                    break;
                case SetupArgumentType.IsAny:
                    itMemberInfo = Reflection.GetMethodExt(itType, nameof(It.IsAny)).MakeGenericMethod(parameterType);
                    break;
                case SetupArgumentType.IsIn:
                    itMemberInfo = Reflection.GetMethodExt(itType, nameof(It.IsIn), typeof(IEnumerable<>)).MakeGenericMethod(parameterType);
                    var arrayInit = Expression.NewArrayInit(parameterType,
                        definition.Values.Select(x => Expression.Convert(Expression.Constant(TryCast(x, parameterType, out var c) ? c : x), parameterType)));
                    itArgumentExpressions.Add(arrayInit);
                    break;
                case SetupArgumentType.IsInRange:
                    itMemberInfo = Reflection.GetMethodExt(itType, nameof(It.IsInRange), typeof(MethodReflection.T),
                        typeof(MethodReflection.T), typeof(Range)).MakeGenericMethod(parameterType);
                    itArgumentExpressions.Add(Expression.Convert(Expression.Constant(TryCast(definition.Values[0], parameterType, out var from) ? from : definition.Values[0]), parameterType));
                    itArgumentExpressions.Add(Expression.Convert(Expression.Constant(TryCast(definition.Values[1], parameterType, out var to) ? to : definition.Values[1]), parameterType));
                    itArgumentExpressions.Add(definition.Range != null
                        ? Expression.Constant((Range)definition.Range)
                        : Expression.Constant(Range.Inclusive));
                    break;
                case SetupArgumentType.IsNotIn:
                    itMemberInfo = Reflection.GetMethodExt(itType, nameof(It.IsNotIn), typeof(IEnumerable<>)).MakeGenericMethod(parameterType);
                    var notArrayInit = Expression.NewArrayInit(parameterType,
                        definition.Values.Select(x => Expression.Convert(Expression.Constant(TryCast(x, parameterType, out var c) ? c : x), parameterType)));
                    itArgumentExpressions.Add(notArrayInit);
                    break;
                case SetupArgumentType.IsNotNull:
                    itMemberInfo = Reflection.GetMethodExt(itType, nameof(It.IsNotNull)).MakeGenericMethod(parameterType);
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }

            return Expression.Call(itMemberInfo, itArgumentExpressions);
        }

        private Expression BuildPassByRefArgumentExpression(int index, SetupArgumentDefinition definition, Type refType, Type elementType, ref Dictionary<ParameterExpression, object> forwardedArgs)
        {
            Expression parameterExpression;
            switch (definition.Type)
            {
                case SetupArgumentType.Is:
                    /* Example of how to call a mock w/ ref parameter
                        public void Test()
                        {
                            string o = "1";
                            var mock = new Mock<Hola>();            

                            mock.Setup(x => x.DoSomething(ref o))
                                .Callback(new DoSomethingAction((ref string a) => a = "2"));

                            mock.Object.DoSomething(ref o);

                            Assert.Equal("2", o)
                        }
                     */
                    // TODO: Create a collection of variables with constant assignments to put in the setup expression?
                    // TODO: or better yet, try and pass the definition's value directly as a ref? 
                    // TODO: need to take care that the args passed into DynamicInvoke do not need to be ref'd - it should be
                    // TODO: passed in as values then made ref within the expression tree.
                    var name = $"p{index:00}";
                    var itByRef = ItByRefMemberInfos.Is(elementType).Invoke(null, new [] {definition.Values[0]});
                    var forwardedArgExpression = Expression.Parameter(itByRef.GetType(), name);
                    forwardedArgs.Add(forwardedArgExpression, itByRef);
                    parameterExpression = Expression.Field(forwardedArgExpression, ItByRefMemberInfos.Value(elementType));
                    return parameterExpression;
                case SetupArgumentType.IsAny:
                    var itRefType = typeof(It.Ref<>).MakeGenericType(elementType);
                    var itFieldInfo = itRefType.GetField(nameof(It.IsAny));
                    parameterExpression = Expression.Parameter(itRefType, "r");
                    return Expression.Field(parameterExpression, itFieldInfo);
                case SetupArgumentType.IsIn:
                    throw new NotSupportedException($"The {nameof(SetupArgumentType.IsIn)} type is not implemented for ref argument");
                case SetupArgumentType.IsInRange:
                    throw new NotSupportedException($"The {nameof(SetupArgumentType.IsInRange)} type is not implemented for ref argument");
                case SetupArgumentType.IsNotIn:
                    throw new NotSupportedException($"The {nameof(SetupArgumentType.IsNotIn)} type is not implemented for ref argument");
                case SetupArgumentType.IsNotNull:
                    throw new NotSupportedException($"The {nameof(SetupArgumentType.IsNotIn)} type is not implemented for ref argument");
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        private static (Type type, bool isRef, bool isOut) GetParameterType(ParameterInfo parameterInfo)
        {
            var isRef = false;
            var isOut = false;
            var parameterType = parameterInfo.ParameterType;

            if (parameterType.IsByRef && parameterInfo.IsOut)
            {
                isOut = true;
            }

            if (!parameterType.IsByRef && !parameterInfo.IsOut)
            {
                return (parameterType, isRef, isOut);
            }

            isRef = true;
            parameterType = parameterType.HasElementType ? parameterType.GetElementType() : parameterType;

            return (parameterType, isRef, isOut);
        }

        private static bool TryCast(object value, Type type, out object convertedValue)
        {
            convertedValue = null;

            try
            {
                convertedValue = VariantConverter.ChangeType(value, type);
            }
            catch
            {
                try
                {
                    convertedValue = Convert.ChangeType(value, type);
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
