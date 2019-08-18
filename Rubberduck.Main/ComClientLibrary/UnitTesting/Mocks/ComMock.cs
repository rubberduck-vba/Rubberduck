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
        private readonly SetupArgumentResolver _resolver;

        public ComMock(Mock mock, Type type, IEnumerable<Type> supportedInterfaces)
        {
            Mock = mock;
            _type = type;
            _supportedInterfaces = supportedInterfaces;
            _resolver = new SetupArgumentResolver();

            Mock.As<IComMocked>().Setup(x => x.Mock).Returns(this);
            mocked = new ComMocked(this, _supportedInterfaces);
        }

        /// <remarks>
        /// Refer to remarks in <see cref="SetupArgumentResolver.ResolveArgs"/> for how the
        /// parameter <paramref name="Args"/> is handled. 
        /// </remarks>
        public void SetupWithReturns(string Name, object Value, object Args = null)
        {
            var args = _resolver.ResolveArgs(Args);
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

        /// <remarks>
        /// Refer to remarks in <see cref="SetupArgumentResolver.ResolveArgs"/> for how the
        /// parameter <paramref name="Args"/> is handled. 
        /// </remarks>
        public void SetupWithCallback(string Name, Action Callback, object Args = null)
        {
            var args = _resolver.ResolveArgs(Args);
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
        /// Builds the basic Setup expressions using provided inputs. The return can be then expanded upon for
        /// specifying behaviors of the given Setup.
        /// </summary>
        /// <param name="name">The member name on the mocked's interface</param>
        /// <param name="args">Arguments required for the member, if any. If none are required, pass in null</param>
        /// <returns>An list of <see cref="Expression"/> representing the mock's Setup expression ("mock.Setup(...)") for each implemented interface</returns>
        private IEnumerable<SetupData> CreateSetupExpression(string name, SetupArgumentDefinitions args)
        {
            var setupDatas = new List<SetupData>();
            var membersToSetup = GetMembers(name);
            var parameterExpressions = _resolver.ResolveParameters(membersToSetup.Parameters.ToArray(), args);
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