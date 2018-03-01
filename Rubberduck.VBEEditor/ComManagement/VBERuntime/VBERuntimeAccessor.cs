using System;

namespace Rubberduck.VBEditor.ComManagement.VBERuntime
{
    public class VBERuntimeAccessor : IVBERuntime
    {
        private enum DllVersion
        {
            Unknown,
            Vbe6,
            Vbe7
        }

        private static DllVersion _version;
        private readonly IVBERuntime _runtime;
        
        static VBERuntimeAccessor()
        {
            _version = DllVersion.Unknown;
        }

        public VBERuntimeAccessor()
        {
            switch (_version)
            {
                case DllVersion.Vbe7:
                    _runtime = new VBERuntime7();
                    break;
                case DllVersion.Vbe6:
                    _runtime = new VBERuntime6();
                    break;
                default:
                    _runtime = DetermineVersion();
                    break;
            }
        }

        private IVBERuntime DetermineVersion()
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
