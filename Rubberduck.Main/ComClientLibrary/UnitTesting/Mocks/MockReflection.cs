using System;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using ExpressiveReflection;
using Moq;
using Moq.Language;
using Moq.Language.Flow;

namespace Rubberduck.ComClientLibrary.UnitTesting.Mocks
{
    /// <remarks>
    /// Most methods on the <see cref="Mock{T}"/> are generic. Because the <see cref="MethodInfo"/> are
    /// different for each closed generic type, we cannot use the open generic <see cref="MethodInfo"/>.
    /// To address this, we need to get the object, and look up the equivalent <see cref="MethodInfo"/> via
    /// handles which are the same regardless of generic parameters used, and return the "closed" version.
    /// </remarks>
    public static class MockMemberInfos
    {
        public static MethodInfo As(Type type)
        {
            return Reflection.GetMethodExt(typeof(Mock), MockMemberNames.As()).MakeGenericMethod(type);
        }

        public static MethodInfo Verify(Type mockType)
        {
            var typeHandle = mockType.TypeHandle;
            var mock = typeof(Mock<>);

            var actionArgExpression = typeof(Expression<>).MakeGenericType(typeof(Action<>));
            var genericMethod = Reflection.GetMethodExt(mock, MockMemberNames.Verify(), actionArgExpression, typeof(Moq.Times));
            var specificMethod = (MethodInfo)MethodBase.GetMethodFromHandle(genericMethod.MethodHandle, typeHandle);

            return specificMethod;
        }

        public static MethodInfo Verify(Type mockType, Type returnType)
        {
            var typeHandle = mockType.TypeHandle;
            var mock = typeof(Mock<>);

            var funcArgExpression = typeof(Expression<>).MakeGenericType(returnType != null ?
                typeof(Func<,>) :
                typeof(Action<>)
            );
            var genericMethod = Reflection.GetMethodExt(mock, MockMemberNames.Verify(), funcArgExpression, typeof(Moq.Times));
            var specificMethod = (MethodInfo)MethodBase.GetMethodFromHandle(genericMethod.MethodHandle, typeHandle);

            return returnType != null ? specificMethod.MakeGenericMethod(returnType) : specificMethod;
        }

        public static MethodInfo Setup(Type mockType, Type returnType)
        {
            var typeHandle = mockType.TypeHandle;
            var mock = typeof(Mock<>);
            var expression = typeof(Expression<>).MakeGenericType(returnType != null ?  
                typeof(Func<,>) :
                typeof(Action<>)
            );
            var genericMethod = Reflection.GetMethodExt(mock, MockMemberNames.Setup(), expression);
            var specificMethod = (MethodInfo) MethodBase.GetMethodFromHandle(genericMethod.MethodHandle, typeHandle);
            return returnType != null ? specificMethod.MakeGenericMethod(returnType) : specificMethod;
        }

        public static MethodInfo Returns(Type setupMockType)
        {
            var typeHandle = setupMockType.GetInterfaces().Single(i =>
                i.IsGenericType &&
                i.GetGenericTypeDefinition() == typeof(IReturns<,>)
            ).TypeHandle;
            var setup = typeof(IReturns<,>);
            var result = typeof(MethodReflection.T); 
            var genericMethod = Reflection.GetMethodExt(setup, MockMemberNames.Returns(), result);
            return (MethodInfo) MethodBase.GetMethodFromHandle(genericMethod.MethodHandle, typeHandle);
        }

        public static MethodInfo Callback(Type setupMockType)
        {
            var typeHandle = setupMockType.GetInterfaces().Single(i =>
                !i.IsGenericType &&
                i == typeof(ICallback)
            ).TypeHandle;
            var setup = typeof(ICallback);
            var callback = typeof(Delegate);
            var genericMethod = Reflection.GetMethodExt(setup, MockMemberNames.Callback(), callback);
            return (MethodInfo) MethodBase.GetMethodFromHandle(genericMethod.MethodHandle, typeHandle);
        }
    }

    /// <remarks>
    /// Though most members are generic, for the purposes of getting the names
    /// they are all the same regardless of the actual closed generic types used
    /// so we can just use object as a placeholder for the generic parameters.
    /// </remarks>
    public static class MockMemberNames
    {
        public static string As()
        {
            return nameof(Mock.As);
        }

        public static string Setup()
        {
            return nameof(Mock<object>.Setup);
        }

        public static string Returns()
        {
            return nameof(ISetup<object, object>.Returns);
        }

        public static string Callback()
        {
            return nameof(ISetup<object>.Callback);
        }

        public static string Verify()
        {
            return nameof(Mock.Verify);
        }
    }
}
