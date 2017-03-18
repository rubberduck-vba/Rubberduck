using System.ComponentModel;
using System.Runtime.InteropServices;

namespace Rubberduck.UnitTesting
{
    [ComVisible(true)]
    [Guid(RubberduckGuid.IFakesProviderGuid)]
    [EditorBrowsable(EditorBrowsableState.Always)]
    public interface IFakesProvider
    {
        IFake MsgBox { get; }
    }
}
