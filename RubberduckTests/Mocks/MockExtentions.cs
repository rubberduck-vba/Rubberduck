using System.Runtime.CompilerServices;
using Moq;

namespace RubberduckTests.Mocks
{
    public static class MockExtentions
    {
        public static Mock<T> SetupReferenceEqualityIncludingHashCode<T>(this Mock<T> mock) where T : class
        {
            mock.Setup(m => m.Equals(It.IsAny<object>()))
                .Returns((object other) => ReferenceEquals(mock.Object, other));
            mock.Setup(m => m.GetHashCode())
                .Returns(() => RuntimeHelpers.GetHashCode(mock.Object));

            return mock;
        }
    }
}
