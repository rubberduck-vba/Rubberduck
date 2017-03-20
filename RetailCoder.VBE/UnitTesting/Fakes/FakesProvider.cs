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
        internal const int AllInvocations = -1;
        // ReSharper disable once InconsistentNaming
        public const int rdAllInvocations = AllInvocations;     

        private static Dictionary<Type, StubBase> ActiveFakes { get; } = new Dictionary<Type, StubBase>();

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

        private T RetrieveOrCreateFunction<T>(Type type) where T : class
        {
            CodeIsUnderTest = true;
            if (!ActiveFakes.ContainsKey(type))
            {
                ActiveFakes.Add(type, (StubBase)Activator.CreateInstance(type));
            }
            return ActiveFakes[type] as T;
        }

        #region Function Overrides

        public IFake MsgBox => RetrieveOrCreateFunction<IFake>(typeof(MsgBox));
        public IFake InputBox => RetrieveOrCreateFunction<IFake>(typeof(InputBox));
        public IStub Beep => RetrieveOrCreateFunction<IStub>(typeof(Beep));

        #endregion
    }
}
