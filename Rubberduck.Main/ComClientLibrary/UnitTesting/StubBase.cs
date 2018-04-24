using System;
using System.Collections.Generic;
using EasyHook;
using Rubberduck.Parsing.ComReflection;

namespace Rubberduck.UnitTesting
{
    internal class StubBase : IStub, IDisposable
    {
        internal const string TargetLibrary = "vbe7.dll";
        private readonly List<LocalHook> _hooks = new List<LocalHook>();

        #region Internal

        protected void InjectDelegate(Delegate callbackDelegate, IntPtr procAddress)
        {
            var hook = LocalHook.Create(procAddress, callbackDelegate, null);
            hook.ThreadACL.SetInclusiveACL(new[] { 0 });
            _hooks.Add(hook);
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
            foreach (var hook in _hooks)
            {
                hook.Dispose();
            }
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
