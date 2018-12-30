using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.InteropServices;
using Moq;
using Rubberduck.Resources.Registration;

// ReSharper disable InconsistentNaming

namespace Rubberduck.ComClientLibrary.UnitTesting.Mocks
{
    [
        ComVisible(true),
        Guid(RubberduckGuid.ComMockGuid),
        ProgId(RubberduckProgId.ComMockProgId),
        ClassInterface(ClassInterfaceType.None),
        ComDefaultInterface(typeof(IComMock))
    ]
    public class ComMock : IComMock
    {
        private readonly Mock _mock;
        private readonly Type _type;
        private readonly IEnumerable<Type> _supportedTypes;

        public ComMock(Mock mock, Type type, IEnumerable<Type> supportedTypes)
        {
            _mock = mock;
            _type = type;
            _supportedTypes = supportedTypes;
        }

        public void SetupWithReturns(string Name, object Value, object Args = null)
        {
            var args = ResolveArgs(Args);
            var setupData = CreateSetupExpression(Name, args);

            var returnsType = setupData.SetupMemberInfo.ReturnType.GetInterface("IReturns`2");
            var returnsMemberInfos = returnsType.GetMember("Returns");

            //TODO: find a better way to get the correct method
            var returnsMemberInfo = (MethodInfo)returnsMemberInfos.First();

            Debug.Assert(returnsMemberInfo != null);

            var valueParameterExpression = Expression.Parameter(setupData.TargetType, "value");

            var castReturnExpression = Expression.Convert(setupData.SetupExpression, returnsType);
            var returnsCallExpression = Expression.Call(castReturnExpression, returnsMemberInfo, valueParameterExpression);
            var lambda = Expression.Lambda(returnsCallExpression, setupData.MockParameterExpression, valueParameterExpression);
            lambda.Compile().DynamicInvoke(_mock, Value);
        }

        public void SetupWithCallback(string Name, Action Callback, object Args = null)
        {
            var args = ResolveArgs(Args);
            var setupData = CreateSetupExpression(Name, args);

            var callbackType = setupData.SetupMemberInfo.ReturnType.GetInterface("ICallback`2");
            var callbackMemberInfos = callbackType.GetMember("Callback");

            //TODO: find a better way to get the correct method
            var callbackMemberInfo = (MethodInfo) callbackMemberInfos.First();
            
            Debug.Assert(callbackMemberInfo != null);

            var valueParameterExpression = Expression.Parameter(Callback.GetType(), "value");

            var castCallbackExpression = Expression.Convert(setupData.SetupExpression, callbackType);
            var callCallbackExpression =
                Expression.Call(castCallbackExpression, callbackMemberInfo, valueParameterExpression);
            var lambda = Expression.Lambda(callCallbackExpression, setupData.MockParameterExpression,
                valueParameterExpression);
            lambda.Compile().DynamicInvoke(_mock, Callback);
        }

        public object Object => new ComMocked(_mock.Object, _supportedTypes);

        /// <summary>
        /// Converts a variant args into the <see cref="MockArgumentDefinitions"/> collection. This supports calls from COM
        /// using the Variant data type.
        /// </summary>
        /// <remarks>
        /// The procedure needs to handle the following cases where the variant...:
        ///   1) contains a single value
        ///   2) contains an Array() of values
        ///   3) wraps a single <see cref="MockArgumentDefinition"/>
        ///   4) points to a <see cref="MockArgumentDefinitions"/> collection.
        ///   5) wraps a single <see cref="Missing"/> object in which case we return a null
        ///   6) wraps an array of single <see cref="Missing"/> object in which case we return a null
        /// </remarks>
        /// <param name="args">Should be a COM Variant that can be cast into valid values as explained in the remarks</param>
        /// <returns>A <see cref="MockArgumentDefinitions"/> collection or null</returns>
        private static MockArgumentDefinitions ResolveArgs(object args)
        {
            // We must ensure that the arrays are resolved before calling the
            // single object wrapper to ensure we don't end up wrapping the 
            // arrays as a single value; do not change the switch order willy-nilly.
            switch (args)
            {
                case Missing missing:
                    return null;
                case MockArgumentDefinitions definitions:
                    return definitions;
                case MockArgumentDefinition definition:
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

        private static MockArgumentDefinitions WrapArgumentDefinitions(object singleObject)
        {
            var list = new MockArgumentDefinitions();
            var inDefinition = new MockArgumentDefinition(MockArgumentType.Is, new[] {singleObject});
            list.Add(inDefinition);
            return list;
        }

        private static MockArgumentDefinitions WrapArgumentDefinitions(object[] objects)
        {
            var list = new MockArgumentDefinitions();
            foreach (var item in objects)
            {
                switch (item)
                {
                    case MockArgumentDefinition argumentDefinition:
                        list.Add(argumentDefinition);
                        break;
                    case object[] arrayObjects:
                        var inDefinition = new MockArgumentDefinition(MockArgumentType.IsIn, arrayObjects);
                        list.Add(inDefinition);
                        break;
                    case object singleObject:
                        var isDefinition =
                            new MockArgumentDefinition(MockArgumentType.Is, new[] {singleObject});
                        list.Add(isDefinition);
                        break;
                }
            }

            return list;
        }

        private static MockArgumentDefinitions WrapArgumentDefinitions(MockArgumentDefinition mockArgumentDefinition)
        {
            return new MockArgumentDefinitions
            {
                mockArgumentDefinition
            };
        }

        /// <summary>
        /// Transform the collection of <see cref="MockArgumentDefinition"/> into a <see cref="IReadOnlyList{Expression}"/>
        /// </summary>
        /// <remarks>
        /// If a method `Foo` requires one argument, we need to specify the behavior in an expression similar
        /// to this: `Mock<>.Setup(x => x.Foo(It.IsAny<>())`. The class <see cref="It"/> is static so we can
        /// create call expresssions directly on it. 
        /// </remarks>
        /// <param name="parameters">Array of <see cref="ParameterInfo"/> returned from the member for which the <see cref="MockArgumentDefinitions"/> applies to</param>
        /// <param name="args">The <see cref="MockArgumentDefinitions"/> collection containing user supplied behavior</param>
        /// <returns>A read-only list containing the <see cref="Expression"/> of arguments</returns>
        private IReadOnlyList<Expression> ResolveParameters(ParameterInfo[] parameters,
            MockArgumentDefinitions args)
        {
            var argsCount = args?.Count ?? 0;
            if (parameters.Length != argsCount)
            {
                throw new ArgumentOutOfRangeException(nameof(args),
                    $"The method expects {parameters.Length} parameters but only {argsCount} argument definitions were supplied. Setting up the mock's behavior requires that all parameters be filled in.");
            }
            if (parameters.Length == 0)
            {
                return null;
            }

            var resolvedArguments = new List<Expression>();
            for(var i = 0; i < parameters.Length; i ++)
            {
                var parameter = parameters[i];
                var definition = args.Item(i);

                var itType = typeof(It);
                MethodInfo itMemberInfo;
                
                var parameterType = parameter.ParameterType;
                var itArgumentExpressions = new List<Expression>();

                switch (definition.Type)
                {
                    case MockArgumentType.Is:
                        itMemberInfo = itType.GetMethod(nameof(It.Is)).MakeGenericMethod(parameterType);
                        itArgumentExpressions.Add(Expression.Constant(definition.Values[0]));
                        break;
                    case MockArgumentType.IsAny:
                        itMemberInfo = itType.GetMethod(nameof(It.IsAny)).MakeGenericMethod(parameterType);
                        break;
                    case MockArgumentType.IsIn:
                        itMemberInfo = itType.GetMethod(nameof(It.IsIn)).MakeGenericMethod(parameterType);
                        itArgumentExpressions.AddRange(definition.Values.Select(val => Expression.Constant(val)));
                        break;
                    case MockArgumentType.IsInRange:
                        itMemberInfo = itType.GetMethod(nameof(It.IsInRange)).MakeGenericMethod(parameterType);
                        itArgumentExpressions.Add(Expression.Constant(definition.Values[0]));
                        itArgumentExpressions.Add(Expression.Constant(definition.Values[1]));
                        if (definition.Range != null)
                        {
                            itArgumentExpressions.Add(Expression.Constant((Range)definition.Range));
                        }
                        break;
                    case MockArgumentType.IsNotIn:
                        itMemberInfo = itType.GetMethod(nameof(It.IsNotIn)).MakeGenericMethod(parameterType);
                        itArgumentExpressions.AddRange(definition.Values.Select(val => Expression.Constant(val)));
                        break;
                    case MockArgumentType.IsNotNull:
                        itMemberInfo = itType.GetMethod(nameof(It.IsNotNull)).MakeGenericMethod(parameterType);
                        break;
                    default:
                        throw new ArgumentOutOfRangeException();
                }

                var callExpression = Expression.Call(itMemberInfo, itArgumentExpressions);
                resolvedArguments.Add(callExpression);
            }

            return resolvedArguments;
        }

        /// <summary>
        /// Builds the basic Setup expression using provided inputs. 
        /// </summary>
        /// <param name="name">The member name on the mocked's interface</param>
        /// <param name="args">Arguments required for the member, if any. If none are required, pass in null</param>
        /// <returns>An <see cref="Expression"/> representing the mock's Setup expression ("mock.Setup(...)")</returns>
        private SetupData CreateSetupExpression(string name, MockArgumentDefinitions args)
        {
            // This represents the left hand side of the lambda expression within the `Setup()`
            var memberType = GetMockMemberTypeByName(name);
            var typeExpression = Expression.Parameter(_type, "x");

            // We need to know whether we need to to make a expression `x.SomeProperty` or `x.SomeMethod()`
            var memberInfo = GetTypeMemberInfo(name);
            Expression memberAccessExpression;

            switch (memberInfo)
            {
                case FieldInfo _:
                    memberAccessExpression = Expression.MakeMemberAccess(typeExpression, memberInfo);
                    break;
                case PropertyInfo propertyInfo:
                    var propertyParameterExpressions = ResolveParameters(propertyInfo.GetIndexParameters(), args);
                    memberAccessExpression = propertyParameterExpressions != null
                        ? (Expression)Expression.Property(typeExpression, propertyInfo, propertyParameterExpressions)
                        : Expression.MakeMemberAccess(typeExpression, memberInfo);
                    break;
                case MethodInfo methodInfo:
                    var methodParameterExpressions = ResolveParameters(methodInfo.GetParameters(), args);
                    memberAccessExpression = methodParameterExpressions != null
                        ? Expression.Call(typeExpression, methodInfo, methodParameterExpressions)
                        : Expression.Call(typeExpression, methodInfo);
                    break;
                default:
                    throw new InvalidCastException($"Could not convert the member to a field, a property or a method");
            }
            
            // Finalize the expression within the Setup's lambda. 
            var lambdaTypes = Expression.GetFuncType(_type, memberType);
            var expression = Expression.Lambda(lambdaTypes, memberAccessExpression, typeExpression);

            // In order to be able to invoke the Setup method, we must have a 
            // generic Mock and we only have the non-generic mock. Thus, we must
            // do a `_mock.As<_type>()` first.
            var mockType = _mock.GetType();
            var asMemberInfo = mockType.GetMethod("As")?.MakeGenericMethod(_type);

            
            var setupMemberInfos = mockType.GetMember("Setup");

            //TODO: Find a better way to find the correct type of Setup method
            var setupMemberInfo = ((MethodInfo)setupMemberInfos.Last()).MakeGenericMethod(memberType);

            Debug.Assert(asMemberInfo != null);
            Debug.Assert(setupMemberInfo != null);

            // Create the expression for invoking `Mock<>.Setup()`, providing the expression above.
            // We also want the _mock to be input into this lambda, so we must also make the `mock`
            // a parameter so that we can pass the _mock in when we dynamically invoke the expression.
            var mockParameterExpression = Expression.Parameter(mockType, "mock");
            var asCallExpression = Expression.Call(mockParameterExpression, asMemberInfo);
            var setupCallExpression = Expression.Call(asCallExpression, setupMemberInfo, expression);

            // At this point, we have the expression for the `Mock<>.Setup().`. The expression can be
            // further processed for additional actions (e.g. adding `Returns()` or `Callback()`. 
            return new SetupData(setupCallExpression, setupMemberInfo, memberType, mockParameterExpression, args);
        }
        
        private MemberInfo GetTypeMemberInfo(string name)
        {
            // TODO This currently searches the only default interface. However, a COM type
            // can have multiple interface and we already know about them via the _supportedTypes
            // member. We should be able to enumerate all the types until we find. However, we 
            // need to consider where one interface hides the member over the other. 
            var memberInfos = _type.GetMember(name);

            // COM does not allow overloading of members, so we should not expect more than one version to exist
            // However, we might have no matches because the name parameter was malformed or mispelt, which 
            // should throw at that point. 
            Debug.Assert(memberInfos.Length >= 0 && memberInfos.Length <= 1);

            if (memberInfos.Length == 0)
            {
                throw new ArgumentOutOfRangeException(name, $"The member was not found on the interface '{_type.Name}'");
            }

            return memberInfos.First();
        }

        private Type GetMockMemberTypeByName(string name)
        {
            var members = _type.GetMember(name);

            // Because COM does not allow for overloading members, we should not have more than one entry for a given name.
            Debug.Assert(members.Length == 1);

            var memberInfo = members.FirstOrDefault();

            if (memberInfo == null)
            {
                throw new ArgumentOutOfRangeException(name, $"Not found on the interface '{_type.Name}'");
            }

            switch (memberInfo)
            {
                case PropertyInfo propertyInfo:
                    return propertyInfo.PropertyType;
                case MethodInfo methodInfo:
                    return methodInfo.ReturnType;
                case FieldInfo fieldInfo:
                    return fieldInfo.FieldType;
                default:
                    throw new ArgumentOutOfRangeException(name, $"Found on the interface '{_type.Name}' but seems to be neither a method nor a property nor a field; the member info type was {memberInfo.GetType()}");
            }
        }

        /// <summary>
        /// Provides base for building a <see cref="Mock{T}.Setup"/> lambda, returned by <see cref="ComMock.CreateSetupExpression"/>
        /// This is used for further developing the lambda expression to invoke other methods that would be provided by the result of
        /// the Setup().
        /// </summary>
        private readonly struct SetupData
        {
            /// <summary>
            /// The base expression representing <see cref="Mock{T}.Setup"/>. Refer to <see cref="ComMock.CreateSetupExpression"/> for details.
            /// This is usually used as a start for further development of the lambda.
            /// </summary>
            internal Expression SetupExpression { get; }

            /// <summary>
            /// MethodInfo on <see cref="Mock{T}.Setup"/>. Useful for supporting casts to different interfaces to expose
            /// methods from those interfaces. 
            /// </summary>
            internal MethodInfo SetupMemberInfo { get; }

            /// <summary>
            /// The type of the member of mocked type, determined at runtime.
            /// </summary>
            internal Type TargetType { get; }

            /// <summary>
            /// The parameter expression for the mock. Required for the final lambda construction because we must pass in
            /// the <see cref="ComMock._mock"/> to the lambda, usually like this: `(mock) => mock.As()`.
            /// </summary>
            internal ParameterExpression MockParameterExpression { get; }

            /// <summary>
            /// The mock argument definitions contains the user-supplied values. Required for the final lambda construction.
            /// Should be equivalent to `(mock, arg1, arg2) => mock.As<>.Setup(someMethod(arg1, arg2)...`. The values should
            /// created using <see cref="Moq.It"/> class to ensure correct behavior.
            /// </summary>
            internal MockArgumentDefinitions MockArgumentDefinitions { get; }

            internal SetupData(Expression setupExpression, MethodInfo setupMemberInfo, Type targetType, ParameterExpression mockParameterExpression, MockArgumentDefinitions mockArgumentDefinitions)
            {
                SetupExpression = setupExpression;
                SetupMemberInfo = setupMemberInfo;
                TargetType = targetType;
                MockParameterExpression = mockParameterExpression;
                MockArgumentDefinitions = mockArgumentDefinitions;
            }
        }
    }
}