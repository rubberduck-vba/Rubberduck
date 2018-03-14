using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Rubberduck.UI;

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
                        string.Format(RubberduckUI.Assert_FailedMessageFormat, methodName, message).Trim()));
        }

        public static void OnAssertInconclusive(string message)
        {
            OnAssertCompleted?.Invoke(null, new AssertCompletedEventArgs(TestOutcome.Inconclusive, message));
        }

        public static void OnAssertIgnored()
        {
            OnAssertCompleted?.Invoke(null, new AssertCompletedEventArgs(TestOutcome.Ignored));
        }

        [DllImport("vbe7.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.IUnknown)]
        private static extern object rtcErrObj();

        public static void RaiseVbaError(int number, string source = "", string description = "", string helpfile = "", int helpcontext = 0)
        {
            OnAssertInconclusive(RubberduckUI.Assert_NotImplemented);
        }
    }
}
