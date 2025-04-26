using Rubberduck.Resources.Registration;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.ComManagement.TypeLibs;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

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
        void InitializeAPIs(ITestEngine testEngine);
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
        private readonly IVBETypeLibsAPI_Object _vbeTypeLibsAPI_Object;
        private ITestEngineAPI _testEngineAPI;

        public ExternalAPI(IVBETypeLibsAPI_Object vbeTypeLibsAPI_Object)
        {
            _vbeTypeLibsAPI_Object = vbeTypeLibsAPI_Object;
        }

        public void InitializeAPIs(ITestEngine testEngine)
        {
            _testEngineAPI = new TestEngineAPI(testEngine);
        }

        public IVBETypeLibsAPI_Object VBETypeLibsAPI { get => _vbeTypeLibsAPI_Object; }
        public ITestEngineAPI TestEngineAPI
        {
            get
            {
                if (_testEngineAPI == null)
                {
                    throw new InvalidOperationException("TestEngineAPI is not initialized.");
                }
                return _testEngineAPI;
            }
        }
    }
}
