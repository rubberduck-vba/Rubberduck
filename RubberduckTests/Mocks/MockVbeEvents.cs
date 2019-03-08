using Moq;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.Mocks
{
    public class MockVbeEvents
    {
        public static Mock<IVbeEvents> CreateMockVbeEvents(Mock<IVBE> vbe)
        {
            var result = new Mock<IVbeEvents>();
            result.SetupReferenceEqualityIncludingHashCode();
            return result;
        }
    }
}
