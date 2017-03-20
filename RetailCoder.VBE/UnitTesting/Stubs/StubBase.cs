using System;
using Rubberduck.Parsing.ComReflection;

namespace Rubberduck.UnitTesting
{
    internal class StubBase : IStub, IDisposable
    {
        internal const string TargetLibrary = "vbe7.dll";
        private readonly IntPtr _procAddress;
        private EasyHook.LocalHook _hook;

        #region Internal

        internal StubBase(IntPtr procAddress)
        {
            _procAddress = procAddress;
        }

        protected void InjectDelegate(Delegate callbackDelegate)
        {
            _hook = EasyHook.LocalHook.Create(_procAddress, callbackDelegate, null);
            _hook.ThreadACL.SetInclusiveACL(new[] { 0 });
        }

        protected Verifier Verifier { get; } = new Verifier();
        internal uint InvocationCount { get; set; }
        internal bool Throws { get; set; }
        internal string ErrorDescription { get; set; }
        internal int ErrorNumber { get; set; }

        protected void TrackUsage(string parameter, IntPtr value)
        {
            // TODO: Resolve TypeName.
            var variant = new ComVariant(value);
            TrackUsage(parameter, variant, string.Empty);
        }

        protected virtual void TrackUsage(string parameter, object value, string typeName)
        {
            Verifier.AddUsage(parameter, value, typeName, InvocationCount);
        }

        protected void OnCallBack(bool trackNoParams = false)
        {
            InvocationCount++;

            if (trackNoParams)
            {
                Verifier.AddUsage(string.Empty, null, string.Empty, InvocationCount);
            }

            if (Throws)
            {
                AssertHandler.RaiseVbaError(ErrorNumber, ErrorDescription);
            }
        }

        public virtual void Dispose()
        {
            _hook.Dispose();
        }

        #endregion

        #region IStub

        public IVerify Verify => Verifier;

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

        public virtual bool PassThrough { get; set; }

        #endregion
    }
}
