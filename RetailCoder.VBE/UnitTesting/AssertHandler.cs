using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.UnitTesting
{
    internal static class AssertHandler
    {
        internal static event EventHandler<AssertCompletedEventArgs> OnAssertCompleted;

        public static void OnAssertSucceeded()
        {
            var handler = OnAssertCompleted;
            if (handler != null)
            {
                handler(null, new AssertCompletedEventArgs(TestResult.Success()));
            }
        }

        public static void OnAssertFailed(string methodName, string message)
        {
            var handler = OnAssertCompleted;
            if (handler != null)
            {
                handler(null, new AssertCompletedEventArgs(
                                new TestResult(TestOutcome.Failed,
                                methodName + " assertion failed." + (string.IsNullOrEmpty(message) ? string.Empty : " " + message))));
            }
        }

        public static void OnAssertInconclusive(string message)
        {
            var handler = OnAssertCompleted;
            if (handler != null)
            {
                handler(null, new AssertCompletedEventArgs(TestResult.Inconclusive(message)));
            }
        }
    }
}
