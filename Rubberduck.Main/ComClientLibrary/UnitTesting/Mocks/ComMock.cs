using Moq;
using Rubberduck.Resources.Registration;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.InteropServices;

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
        private readonly Type _type;
        private readonly IEnumerable<Type> _supportedInterfaces;
        private readonly ComMocked mocked;

        public ComMock(Mock mock, Type type, IEnumerable<Type> supportedInterfaces)
        {
            Mock = mock;
            _type = type;
            _supportedInterfaces = supportedInterfaces;

            Mock.As<IComMocked>().Setup(x => x.Mock).Returns(this);
            mocked = new ComMocked(this, _supportedInterfaces);
        }

        public void SetupWithReturns(string Name, object Value, object Args = null)
        {
            var args = ResolveArgs(Args);
            var setupDatas = CreateSetupExpression(Name, args);

            foreach (var setupData in setupDatas)
            {
                var builder = MockExpressionBuilder.Create(Mock);
                builder.As(setupData.DeclaringType)
                    .Setup(setupData.SetupExpression, setupData.ReturnType)
                    .Returns(Value, setupData.ReturnType)
                    .Execute();
            }
        }

        public void SetupWithCallback(string Name, Action Callback, object Args = null)
        {
            var args = ResolveArgs(Args);
            var setupDatas = CreateSetupExpression(Name, args);

            foreach (var setupData in setupDatas)
            {
                var builder = MockExpressionBuilder.Create(Mock);
                builder.As(setupData.DeclaringType)
                    .Setup(setupData.SetupExpression)
                    .Callback(Callback)
                    .Execute();
            }
        }

        public object Object => mocked;

        internal Mock Mock { get; }

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
        /// create call expressions directly on it. 
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
                var typeExpression = Expression.Parameter(parameterType, "x");
                
                switch (definition.Type)
                {
                    case MockArgumentType.Is:
                        itMemberInfo = itType.GetMethod(nameof(It.Is)).MakeGenericMethod(parameterType);
                        var bodyExpression = Expression.Equal(typeExpression, Expression.Constant(definition.Values[0]));
                        var itLambda = Expression.Lambda(bodyExpression, typeExpression);
                        itArgumentExpressions.Add(itLambda);
                        break;
                    case MockArgumentType.IsAny:
                        itMemberInfo = itType.GetMethod(nameof(It.IsAny)).MakeGenericMethod(parameterType);
                        break;
                    case MockArgumentType.IsIn:
                        itMemberInfo = ((MethodInfo)itType.GetMember(nameof(It.IsIn)).First()).MakeGenericMethod(parameterType);
                        var arrayInit = Expression.NewArrayInit(parameterType,
                            definition.Values.Select(val => Expression.Constant(val)));
                        itArgumentExpressions.Add(arrayInit);
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
                        itMemberInfo = ((MethodInfo)itType.GetMember(nameof(It.IsNotIn)).First()).MakeGenericMethod(parameterType);
                        var notArrayInit = Expression.NewArrayInit(parameterType,
                            definition.Values.Select(val => Expression.Constant(val)));
                        itArgumentExpressions.Add(notArrayInit);
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
        /// Builds the basic Setup expressions using provided inputs. The return can be then expanded upon for
        /// specifying behaviors of the given Setup.
        /// </summary>
        /// <param name="name">The member name on the mocked's interface</param>
        /// <param name="args">Arguments required for the member, if any. If none are required, pass in null</param>
        /// <returns>An list of <see cref="Expression"/> representing the mock's Setup expression ("mock.Setup(...)") for each implemented interface</returns>
        private IEnumerable<SetupData> CreateSetupExpression(string name, MockArgumentDefinitions args)
        {
            var setupDatas = new List<SetupData>();
            var membersToSetup = GetMembers(name);
            var parameterExpressions = ResolveParameters(membersToSetup.Parameters.ToArray(), args);
            var memberType = membersToSetup.ReturnType;

            foreach (var member in membersToSetup.MemberInfos)
            {
                var typeExpression = Expression.Parameter(member.Key, "x");
                Expression memberAccessExpression;

                switch (member.Value)
                {
                    case FieldInfo _:
                        memberAccessExpression = Expression.MakeMemberAccess(typeExpression, member.Value);
                        break;
                    case PropertyInfo propertyInfo:
                        memberAccessExpression = parameterExpressions != null
                            ? (Expression)Expression.Property(typeExpression, propertyInfo, parameterExpressions)
                            : Expression.MakeMemberAccess(typeExpression, member.Value);
                        break;
                    case MethodInfo methodInfo:
                        memberAccessExpression = parameterExpressions != null
                            ? Expression.Call(typeExpression, methodInfo, parameterExpressions)
                            : Expression.Call(typeExpression, methodInfo);
                        break;
                    default:
                        throw new InvalidCastException($"Could not convert the member to a field, a property or a method");
                }

                // Finalize the expression within the Setup's lambda. 
                var expression = Expression.Lambda(memberAccessExpression, typeExpression);
                
                setupDatas.Add(new SetupData(expression, member.Key, memberType));
            }

            return setupDatas;
        }

        /// <summary>
        /// Discover all members from all interfaces that are named the same and shares the same
        /// signature. 
        /// </summary>
        /// <param name="name">The name of member to find on any interfaces</param>
        /// <returns>
        /// A <see cref="MemberSetupData"/> struct that contains the data for each member needed
        /// to create an expression. See <see cref="CreateSetupExpression"/> for more details.
        /// </returns>
        private MemberSetupData GetMembers(string name)
        {
            var memberInfos = new Dictionary<Type, MemberInfo>();
            Type returnType = null;
            ParameterInfo[] parameters = null;
            MemberInfo member = null;
            var members = _type.GetMember(name);

            //COM interfaces should not allow for method overloading within same interface
            Debug.Assert(members.Length <= 1);

            if (members.Length == 1)
            {
                member = members.First();
                memberInfos.Add(_type, member);

                (returnType, parameters) = GetMemberInfo(member);
            }

            foreach (var subType in _supportedInterfaces)
            {
                if (subType == _type)
                {
                    continue;
                }

                members = subType.GetMember(name);

                //COM interfaces should not allow for method overloading within same interface
                Debug.Assert(members.Length <= 1);

                if (members.Length == 0)
                {
                    continue;
                }

                if (member == null)
                {
                    member = members.First();
                    memberInfos.Add(subType, member);

                    (returnType, parameters) = GetMemberInfo(member);
                }
                else
                {
                    var subMember = members.First();
                    var (subReturnType, subParameters) = GetMemberInfo(member);
                    
                    if (subMember.Name == member.Name &&
                        subMember.MemberType == member.MemberType &&
                        returnType == subReturnType &&
                        parameters.Length == subParameters.Length &&
                        parameters.All(p => subParameters.Any(sp =>
                            p.Name == sp.Name &&
                            p.Position == sp.Position &&
                            p.ParameterType == sp.ParameterType)))
                    {
                        memberInfos.Add(subType, subMember);
                    }
                }
            }

            return new MemberSetupData(memberInfos, returnType, parameters);
        }

        private static (Type returnType, ParameterInfo[] parameters) GetMemberInfo(MemberInfo member)
        {
            Type returnType;
            ParameterInfo[] parameters;

            switch (member)
            {
                case FieldInfo fieldInfo:
                    returnType = fieldInfo.FieldType;
                    parameters = new ParameterInfo[0];
                    break;
                case PropertyInfo propertyInfo:
                    returnType = propertyInfo.PropertyType;
                    parameters = propertyInfo.GetIndexParameters();
                    break;
                case MethodInfo methodInfo:
                    returnType = methodInfo.ReturnType;
                    parameters = methodInfo.GetParameters();
                    break;
                default:
                    throw new ArgumentOutOfRangeException(member.Name, $"Found on the interface '{member.ReflectedType?.Name}' but seems to be neither a method nor a property nor a field; the member info type was {member.GetType()}");
            }

            return (returnType, parameters);
        }

        private readonly struct MemberSetupData
        {
            internal IDictionary<Type, MemberInfo> MemberInfos { get; }
            internal Type ReturnType { get; }
            internal IEnumerable<ParameterInfo> Parameters { get; }

            internal MemberSetupData(IDictionary<Type, MemberInfo> memberInfos, Type returnType,
                IEnumerable<ParameterInfo> parameters)
            {
                MemberInfos = memberInfos;
                ReturnType = returnType;
                Parameters = parameters;
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
            /// The containing interface that implements the member being called in the setup expression.
            /// </summary>
            internal Type DeclaringType { get; }

            /// <summary>
            /// The return type, if any, for the member being called in the setup expression.
            /// </summary>
            internal Type ReturnType { get; }

            internal SetupData(Expression setupExpression, Type declaringType, Type returnType)
            {
                SetupExpression = setupExpression;
                DeclaringType = declaringType;
                ReturnType = returnType;
            }
        }
    }
}