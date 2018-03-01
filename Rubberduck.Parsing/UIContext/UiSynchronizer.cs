using System;
using System.Threading;
using Rubberduck.VBEditor.ComManagement;

namespace Rubberduck.Parsing.UIContext
{
    public static class UiSynchronizer
    {
        private static readonly ReaderWriterLockSlim Sync;
        private const int NoTimeout = -1;
        private const int DefaultTimeout = NoTimeout;

        static UiSynchronizer()
        {
            Sync = new ReaderWriterLockSlim(LockRecursionPolicy.SupportsRecursion);
        }
        
        public static bool RequestComAccess(Func<bool> func, int timeout = DefaultTimeout)
        {
            if (!Sync.TryEnterReadLock(timeout))
            {
                throw new TimeoutException("Timeout exceeded while waiting to acquire a read lock");
            }

            var result = func.Invoke();
            Sync.ExitReadLock();
            return result;
        }

        public static void RequireExclusiveComAccess(Action func, int timeout = DefaultTimeout)
        {
            for (var i = 0; i < timeout || timeout == NoTimeout; i++)
            {
                ComMessagePumper.PumpMessages();

                if (Sync.TryEnterWriteLock(1))
                {
                    break;
                }
            }

            if (!Sync.IsWriteLockHeld)
            {
                throw new TimeoutException("Timeout exceeded while waiting to acquire a write lock");
            }

            func.Invoke();
            Sync.ExitWriteLock();
        }

        public static bool RequireExclusiveComAccess(Func<bool> func,  int timeout = DefaultTimeout, bool DeferToParse = true)
        {
            for (var i = 0; i < timeout || timeout == NoTimeout; i++)
            {
                ComMessagePumper.PumpMessages();

                if (Sync.TryEnterWriteLock(1))
                {
                    break;
                }
            }

            if (!Sync.IsWriteLockHeld)
            {
                throw new TimeoutException("Timeout exceeded while waiting to acquire a write lock");
            }

            var result = func.Invoke();
            if (!DeferToParse)
            {
                Sync.ExitWriteLock();
            }
            return result;
        }

        public static bool ReleaseExclusiveComAccess()
        {
            if (Sync.IsWriteLockHeld)
            {
                Sync.ExitWriteLock();
                return true;
            }

            return false;
        }
    }
}
