using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace Rubberduck.UnitTesting
{
    [ComVisible(true)]
    [ComDefaultInterface(typeof(IFakesProvider))]
    [ProgId(RubberduckProgId.FakesProviderProgId)]
    [Guid(RubberduckGuid.FakesProviderClassGuid)]
    [EditorBrowsable(EditorBrowsableState.Always)]    
    public class FakesProvider : IFakesProvider
    {
        private static Dictionary<Type, FakeBase> ActiveFakes { get; } = new Dictionary<Type, FakeBase>();

        internal bool CodeIsUnderTest { get; set; }

        internal void StartTest()
        {
            if (CodeIsUnderTest)
            {
                return;
            }
            CodeIsUnderTest = true;
        }

        internal void StopTest()
        {           
            foreach (var fake in ActiveFakes.Values)
            {
                fake.Dispose();
            }
            ActiveFakes.Clear();
            CodeIsUnderTest = false;
        }

        private IFake RetrieveOrCreateFake(Type type)
        {
            CodeIsUnderTest = true;
            if (!ActiveFakes.ContainsKey(type))
            {
                ActiveFakes.Add(type, (FakeBase)Activator.CreateInstance(type));
            }
            return ActiveFakes[type];
        }

        #region Function Overrides

        public IFake MsgBox => RetrieveOrCreateFake(typeof(MsgBox));
        public IFake InputBox => RetrieveOrCreateFake(typeof(InputBox));

        #endregion
    }
}
