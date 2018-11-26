using System;

namespace Rubberduck.VBEditor.ComManagement
{
    public static class ComSafeManager
    {
        private static Lazy<IComSafe> _comSafe = new Lazy<IComSafe>(NewComSafe);

        public static IComSafe GetCurrentComSafe()
        {
            return _comSafe.Value;
        }

        public static void DisposeAndResetComSafe()
        {
            var oldComSafe = _comSafe.Value;
            _comSafe = new Lazy<IComSafe>(NewComSafe);
            oldComSafe.Dispose();
        }

        private static IComSafe NewComSafe()
        {
            return new WeakComSafe();
        }
    }
}
