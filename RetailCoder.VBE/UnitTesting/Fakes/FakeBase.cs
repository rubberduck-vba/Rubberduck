using System;
using System.Runtime.InteropServices;

namespace Rubberduck.UnitTesting
{
    internal abstract class FakeBase : IFake, IDisposable
    {
        internal const string TargetLibrary = "vbe7.dll";
        private readonly IntPtr _procAddress;
        private EasyHook.LocalHook _hook;

        #region Internal

        internal FakeBase(IntPtr procAddress)
        {
            _procAddress = procAddress;
        }

        internal uint InvocationCount { get; set; }
        internal object ReturnValue { get; set; }
        internal bool Throws { get; set; }
        internal string ErrorDescription { get; set; }
        internal int ErrorNumber { get; set; }

        protected void InjectDelegate(Delegate callbackDelegate)
        {
            _hook = EasyHook.LocalHook.Create(_procAddress, callbackDelegate, null);
            _hook.ThreadACL.SetInclusiveACL(new[] { 0 });
        }

        public virtual void Dispose()
        {
            _hook.Dispose();
        }

        protected void OnCallBack()
        {
            InvocationCount++;
            if (Throws)
            {
                AssertHandler.RaiseVbaError(ErrorNumber, ErrorDescription);
            }
        }

        protected void TrackUsage(string parameter, object value)
        {
            // TODO: Resolve TypeName.
            _verifier.AddUsage(parameter, value, string.Empty, InvocationCount);
        }

        #endregion

        #region IFake

        public virtual void AssignsByRef(string parameter, object value)            
        {
            throw new NotImplementedException();
        }

        public void RaisesError(int number = 0, string description = "")
        {
            Throws = number != 0;
            ErrorNumber = number;
            ErrorDescription = description;
        }

        public virtual void Returns(object value)
        {
            ReturnValue = value;
        }

        private readonly Verifier _verifier = new Verifier();
        public IVerify Verify => _verifier;

        #endregion
    }
}
