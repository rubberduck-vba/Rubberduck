using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Rubberduck.Resources.UnitTesting;

namespace Rubberduck.UnitTesting
{
    public static class AssertHandler
    {
        public static event EventHandler<AssertCompletedEventArgs> OnAssertCompleted;

        public static void OnAssertSucceeded()
        {
            OnAssertCompleted?.Invoke(null, new AssertCompletedEventArgs(TestOutcome.Succeeded));
        }

        public static void OnAssertFailed(string message, [CallerMemberName] string methodName = "")
        {
            OnAssertCompleted?.Invoke(null,
                    new AssertCompletedEventArgs(TestOutcome.Failed,
                        string.Format(AssertMessages.Assert_FailedMessageFormat, methodName, message).Trim()));
        }

        public static void OnAssertInconclusive(string message)
        {
            OnAssertCompleted?.Invoke(null, new AssertCompletedEventArgs(TestOutcome.Inconclusive, message));
        }


        [DllImport("vbe7.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.IUnknown)]
        private static extern object rtcErrObj();

        public static void RaiseVbaError(int number, string source = "", string description = "", string helpfile = "", int helpcontext = 0)
        {
            OnAssertInconclusive(AssertMessages.Assert_NotImplemented);
        }
    }
}
