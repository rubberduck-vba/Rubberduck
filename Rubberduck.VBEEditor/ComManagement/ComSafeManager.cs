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

        public static void ResetComSafe()
        {
            _comSafe = new Lazy<IComSafe>(NewComSafe);
        }

        private static IComSafe NewComSafe()
        {
            return new WeakComSafe();
        }
    }
}
