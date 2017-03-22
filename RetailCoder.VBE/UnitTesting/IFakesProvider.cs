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
        [DispId(4)]
        IFake Environ { get; }
        [DispId(5)]
        IFake Timer { get; }
        [DispId(6)]
        IFake DoEvents { get; }
        [DispId(7)]
        IFake Shell { get; }
        [DispId(8)]
        IStub SendKeys { get; }        
    }
}
