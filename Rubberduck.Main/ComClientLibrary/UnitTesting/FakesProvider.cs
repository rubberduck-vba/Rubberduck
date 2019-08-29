using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Rubberduck.Resources.Registration;
using Rubberduck.UnitTesting.Fakes;

namespace Rubberduck.UnitTesting
{
    [
        ComVisible(true),
        Guid(RubberduckGuid.FakesProviderClassGuid),
        ProgId(RubberduckProgId.FakesProviderProgId),
        ClassInterface(ClassInterfaceType.None),
        ComDefaultInterface(typeof(IFakesProvider)),
        EditorBrowsable(EditorBrowsableState.Always)
    ]   
    public class FakesProvider : IFakesProvider, IFakes
        // IFakesProvider is COM side, exposed to the VBA User
        // IFakes is Rubberduck side and we inject the FakesProvider back into Core
    {
        internal const int AllInvocations = -1;
        // ReSharper disable once InconsistentNaming - respects COM naming conventions
        [Description("A value indicating that specified configuration applies to all invocations.")]
        public const int rdAllInvocations = AllInvocations;     

        private static Dictionary<Type, StubBase> ActiveFakes { get; } = new Dictionary<Type, StubBase>();

        internal bool CodeIsUnderTest { get; set; }

        public void StartTest()
        {
            if (CodeIsUnderTest)
            {
                return;
            }
            CodeIsUnderTest = true;
        }

        public void StopTest()
        {           
            foreach (var fake in ActiveFakes.Values)
            {
                fake.Dispose();
            }
            ActiveFakes.Clear();
            CodeIsUnderTest = false;
        }

        private T RetrieveOrCreateFunction<T>()
            where T : StubBase, new()
        {
            return RetrieveOrCreateFunction(() => new T());
        }

        private T RetrieveOrCreateFunction<T>(Func<T> factory)
            where T : StubBase
        {
            var type = typeof(T);

            CodeIsUnderTest = true;
            if (!ActiveFakes.ContainsKey(type))
            {
                ActiveFakes.Add(type, factory.Invoke());
            }

            return ActiveFakes[type] as T;
        }

        #region Function Overrides

        public IFake MsgBox => RetrieveOrCreateFunction<MsgBox>();
        public IFake InputBox => RetrieveOrCreateFunction<InputBox>();
        public IStub Beep => RetrieveOrCreateFunction(() => new Beep(VbeProvider.BeepInterceptor));
        public IFake Environ => RetrieveOrCreateFunction<Environ>();
        public IFake Timer => RetrieveOrCreateFunction<Timer>();
        public IFake DoEvents => RetrieveOrCreateFunction<DoEvents>();
        public IFake Shell => RetrieveOrCreateFunction<Shell>();
        public IStub SendKeys => RetrieveOrCreateFunction<SendKeys>();
        public IStub Kill => RetrieveOrCreateFunction<Kill>();
        public IStub MkDir => RetrieveOrCreateFunction<MkDir>();
        public IStub RmDir => RetrieveOrCreateFunction<RmDir>();
        public IStub ChDir => RetrieveOrCreateFunction<ChDir>();
        public IStub ChDrive => RetrieveOrCreateFunction<ChDrive>();
        public IFake CurDir => RetrieveOrCreateFunction<CurDir>();
        public IFake Now => RetrieveOrCreateFunction<Now>();
        public IFake Time => RetrieveOrCreateFunction<Time>();
        public IFake Date => RetrieveOrCreateFunction<Date>();

        #endregion
    }
}
