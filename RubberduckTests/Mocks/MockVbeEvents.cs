using Moq;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.Mocks
{
    public class MockVbeEvents
    {
        public static Mock<IVBEEvents> CreateMockVbeEvents(Mock<IVBE> vbe)
        {
            var result = new Mock<IVBEEvents>();
            result.SetupReferenceEqualityIncludingHashCode();
            return result;
        }
    }
}
