using System.ComponentModel;
using System.Runtime.InteropServices;

namespace Rubberduck.UnitTesting
{
    [ComVisible(true)]
    [Guid(RubberduckGuid.IFakesProviderGuid)]
    [EditorBrowsable(EditorBrowsableState.Always)]
    public interface IFakesProvider
    {
        [DispId(1)]
        IFake MsgBox { get; }
        [DispId(2)]
        IFake InputBox { get; }
        [DispId(3)]
        IStub Beep { get; }
    }
}
