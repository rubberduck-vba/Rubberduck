using Rubberduck.Resources.Registration;
using Rubberduck.VBEditor.ComManagement.TypeLibs;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace Rubberduck.ExternalApi
{
    [
        ComVisible(true),
        Guid(RubberduckGuid.IExternalAPIInterfaceGuid),
        InterfaceType(ComInterfaceType.InterfaceIsDual),
        EditorBrowsable(EditorBrowsableState.Always)
    ]
    public interface IExternalAPI
    {
        [DispId(1)]
        IVBETypeLibsAPI_Object VBETypeLibsAPI { get; }

        [DispId(2)]
        ITestEngineAPI TestEngineAPI { get; }
    }

    [
    ComVisible(true),
    Guid(RubberduckGuid.ExternalAPIObjectGuid),
    ProgId(RubberduckProgId.ExternalAPIObject),
    ClassInterface(ClassInterfaceType.None),
    ComDefaultInterface(typeof(IExternalAPI)),
    EditorBrowsable(EditorBrowsableState.Always)
    ]
    public class ExternalAPI : IExternalAPI
    {
        private readonly Func<IVBETypeLibsAPI_Object> _vbeTypeLibProvider;
        private readonly Func<ITestEngineAPI> _testEngineProvider;

        public ExternalAPI(Func<IVBETypeLibsAPI_Object> vbeTypeLibProvider, Func<ITestEngineAPI> testEngineProvider)
        {
            _vbeTypeLibProvider = vbeTypeLibProvider;
            _testEngineProvider = testEngineProvider;
        }

        public IVBETypeLibsAPI_Object VBETypeLibsAPI => _vbeTypeLibProvider();
        public ITestEngineAPI TestEngineAPI => _testEngineProvider();
    }
}
