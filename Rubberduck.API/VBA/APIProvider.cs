using System.ComponentModel;
using Rubberduck.API.VBA;
using System.Runtime.InteropServices;
using Rubberduck.Resources.Registration;

namespace Rubberduck.API
{
    [
        ComVisible(true),
        Guid(RubberduckGuid.IApiProviderGuid),
        InterfaceType(ComInterfaceType.InterfaceIsDual)
    ]
    public interface IApiProvider
    {
        Parser GetParser(object vbe);
    }

    [
        ComVisible(true),
        Guid(RubberduckGuid.ApiProviderClassGuid),
        ProgId(RubberduckProgId.ApiProviderProgId),
        ClassInterface(ClassInterfaceType.None),
        ComDefaultInterface(typeof(IApiProvider)),
        EditorBrowsable(EditorBrowsableState.Always)
    ]
    public class ApiProvider : IApiProvider
    {
        // vbe is the com coclass interface from the interop assembly.
        // There is no shared interface between VBA and VB6 types, hence object.
        public Parser GetParser(object vbe)
        {
            return  new Parser(vbe);
        }
    }
}
