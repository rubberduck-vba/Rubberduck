using System;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.VBERuntime
{
    public class VBERuntimeAccessor : IVBERuntime
    {
        private static DllVersion _version;
        private readonly IVBERuntime _runtime;
        
        static VBERuntimeAccessor()
        {
            _version = DllVersion.Unknown;
        }
        
        public VBERuntimeAccessor(IVBE vbe)
        {
            if (_version == DllVersion.Unknown)
            {
                try
                {
                    _version = VBEDllVersion.GetCurrentVersion(vbe);
                }
                catch
                {
                    _version = DllVersion.Unknown;
                }
            }
            _runtime = InitializeRuntime();
        }

        private static IVBERuntime InitializeRuntime()
        {
            switch (_version)
            {
                case DllVersion.Vbe7:
                    return new VBERuntime7();
                case DllVersion.Vbe6:
                    return new VBERuntime6();
                default:
                    return DetermineVersion();
            }
        }

        private static IVBERuntime DetermineVersion()
        {
            IVBERuntime runtime;
            try
            {
                runtime = new VBERuntime7();
                runtime.Timer();
                _version = DllVersion.Vbe7;
            }
            catch
            {
                try
                {
                    runtime = new VBERuntime6();
                    runtime.Timer();
                    _version = DllVersion.Vbe6;
                }
                catch
                {
                    // we shouldn't be here.... Rubberduck is a VBA add-in, so how the heck could it have loaded without a VBE dll?!?
                    throw new InvalidOperationException("Cannot execute DoEvents; the VBE dll could not be located.");
                }
            }

            return _version != DllVersion.Unknown ? runtime : null;
        }

        public float Timer()
        {
            return _runtime.Timer();
        }

        public int DoEvents()
        {
            return _runtime.DoEvents();
        }
    }
}
