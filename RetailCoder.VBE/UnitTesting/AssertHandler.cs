using System;
using System.Runtime.CompilerServices;
using Rubberduck.UI;

namespace Rubberduck.UnitTesting
{
    public static class AssertHandler
    {
        public static event EventHandler<AssertCompletedEventArgs> OnAssertCompleted;

        public static void OnAssertSucceeded()
        {
            var handler = OnAssertCompleted;
            if (handler != null)
            {
                handler(null, new AssertCompletedEventArgs(TestOutcome.Succeeded));
            }
        }

        public static void OnAssertFailed(string message, [CallerMemberName] string methodName = "")
        {
            var handler = OnAssertCompleted;
            if (handler != null)
            {
                handler(null,
                    new AssertCompletedEventArgs(TestOutcome.Failed,
                        string.Format(RubberduckUI.Assert_FailedMessageFormat, methodName, message).Trim()));
                                
            }
        }

        public static void OnAssertInconclusive(string message)
        {
            var handler = OnAssertCompleted;
            if (handler != null)
            {
                handler(null, new AssertCompletedEventArgs(TestOutcome.Inconclusive, message));
            }
        }

        public static void OnAssertIgnored()
        {
            var handler = OnAssertCompleted;
            if (handler != null)
            {
                handler(null, new AssertCompletedEventArgs(TestOutcome.Ignored));
            }
        }
    }
}
