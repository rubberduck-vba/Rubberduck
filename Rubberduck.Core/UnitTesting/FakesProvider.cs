using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Rubberduck.UnitTesting.Fakes;

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
        // ReSharper disable once InconsistentNaming - respects COM naming conventions
        [Description("A value indicating that specified configuration applies to all invocations.")]
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
        public IFake Environ => RetrieveOrCreateFunction<IFake>(typeof(Environ));
        public IFake Timer => RetrieveOrCreateFunction<IFake>(typeof(Timer));
        public IFake DoEvents => RetrieveOrCreateFunction<IFake>(typeof(DoEvents));
        public IFake Shell => RetrieveOrCreateFunction<IFake>(typeof(Shell));
        public IStub SendKeys => RetrieveOrCreateFunction<IStub>(typeof(SendKeys));
        public IStub Kill => RetrieveOrCreateFunction<IStub>(typeof(Kill));
        public IStub MkDir => RetrieveOrCreateFunction<IStub>(typeof(MkDir));
        public IStub RmDir => RetrieveOrCreateFunction<IStub>(typeof(RmDir));
        public IStub ChDir => RetrieveOrCreateFunction<IStub>(typeof(ChDir));
        public IStub ChDrive => RetrieveOrCreateFunction<IStub>(typeof(ChDrive));
        //public IFake CurDir => RetrieveOrCreateFunction<IFake>(typeof(CurDir));


        #endregion
    }
}
