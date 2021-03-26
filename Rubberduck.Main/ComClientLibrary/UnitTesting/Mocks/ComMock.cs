using Moq;
using Rubberduck.Resources.Registration;
using System;
using System.Collections.Generic;
using System.Linq;
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
        private readonly ComMocked mocked;
        private readonly SetupArgumentResolver _resolver;
        private readonly SetupExpressionBuilder _setupBuilder;
        private readonly IMockProviderInternal _provider;

        internal ComMock(IMockProviderInternal provider, string project, string progId, Mock mock, Type type, IEnumerable<Type> supportedInterfaces)
        {
            Project = project;
            ProgId = progId;
            Mock = mock;
            _provider = provider;
            _resolver = new SetupArgumentResolver();
            _setupBuilder = new SetupExpressionBuilder(type, supportedInterfaces, _resolver);
            MockedType = type;

            Mock.As<IComMocked>().Setup(x => x.Mock).Returns(this);
            mocked = new ComMocked(this, supportedInterfaces);
        }

        public string Project { get; }

        public string ProgId { get; }

        /// <remarks>
        /// Refer to remarks in <see cref="SetupArgumentResolver.ResolveArgs"/> for how the
        /// parameter <paramref name="Args"/> is handled. 
        /// </remarks>
        public void Setup(string Name, object Args = null)
        {
            var args = _resolver.ResolveArgs(Args);
            var setupDatas = _setupBuilder.CreateExpression(Name, args);

            foreach (var setupData in setupDatas)
            {
                var builder = MockExpressionBuilder.Create(Mock);
                builder.As(setupData.DeclaringType)
                    .Setup(setupData.SetupExpression, setupData.Args)
                    .Execute();
            }
        }

        /// <remarks>
        /// Refer to remarks in <see cref="SetupArgumentResolver.ResolveArgs"/> for how the
        /// parameter <paramref name="Args"/> is handled. 
        /// </remarks>
        public void SetupWithReturns(string Name, object Value, object Args = null)
        {
            var args = _resolver.ResolveArgs(Args);
            var setupDatas = _setupBuilder.CreateExpression(Name, args);

            foreach (var setupData in setupDatas)
            {
                var builder = MockExpressionBuilder.Create(Mock);
                builder.As(setupData.DeclaringType)
                    .Setup(setupData.SetupExpression, setupData.Args, setupData.ReturnType)
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
            var setupDatas = _setupBuilder.CreateExpression(Name, args);

            foreach (var setupData in setupDatas)
            {
                var builder = MockExpressionBuilder.Create(Mock);
                builder.As(setupData.DeclaringType)
                    .Setup(setupData.SetupExpression, setupData.Args)
                    .Callback(Callback)
                    .Execute();
            }
        }

        public IComMock SetupChildMock(string Name, object Args)
        {
            Type type;
            var memberInfo = MockedType.GetMember(Name).FirstOrDefault();
            if (memberInfo == null)
            {
                memberInfo = MockedType.GetInterfaces().SelectMany(face => face.GetMember(Name)).First();
            }

            switch (memberInfo)
            {
                case FieldInfo fieldInfo:
                    type = fieldInfo.FieldType;
                    break;
                case PropertyInfo propertyInfo:
                    type = propertyInfo.PropertyType;
                    break;
                case MethodInfo methodInfo:
                    type = methodInfo.ReturnType;
                    break;
                default:
                    throw new InvalidOperationException($"Couldn't resolve member {Name} and acquire a type to mock.");
            }

            var childMock = _provider.MockChildObject(this, type);
            var target = GetMockedObject(childMock, type);
            SetupWithReturns(Name, target, Args);

            return childMock;
        }

        private object GetMockedObject(IComMock mock, Type type)
        {
            var pUnkSource = IntPtr.Zero;
            var pUnkTarget = IntPtr.Zero;

            try
            {
                pUnkSource = Marshal.GetIUnknownForObject(mock.Object);
                var iid = type.GUID;
                Marshal.QueryInterface(pUnkSource, ref iid, out pUnkTarget);
                return Marshal.GetTypedObjectForIUnknown(pUnkTarget, type);
            }
            finally
            {
                if (pUnkTarget != IntPtr.Zero) Marshal.Release(pUnkTarget);
                if (pUnkSource != IntPtr.Zero) Marshal.Release(pUnkSource);
            }
        }

        public void Verify(string Name, ITimes Times, [MarshalAs(UnmanagedType.Struct), Optional] object Args)
        {
            var args = _resolver.ResolveArgs(Args);
            var setupDatas = _setupBuilder.CreateExpression(Name, args);

            var throwingExecutions = 0;
            MockException lastException = null;
            foreach (var setupData in setupDatas)
            {
                try
                {
                    var builder = MockExpressionBuilder.Create(Mock);
                    builder.As(setupData.DeclaringType)
                        .Verify(setupData.SetupExpression, Times, setupData.Args)
                        .Execute();

                    Rubberduck.UnitTesting.AssertHandler.OnAssertSucceeded();
                }
                catch (TargetInvocationException exception)
                {
                    if (exception.InnerException is MockException inner)
                    {
                        throwingExecutions++;
                        lastException = inner;
                    }
                    else
                    {
                        throw;
                    }
                }
            }
            if (setupDatas.Count() == throwingExecutions)
            {
                // if all mocked interfaces failed the .Verify call, then none of them succeeded:
                Rubberduck.UnitTesting.AssertHandler.OnAssertFailed(lastException.Message);
            }
        }

        public object Object => mocked;

        internal Mock Mock { get; }

        internal Type MockedType { get; }
    }
}