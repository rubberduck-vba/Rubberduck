﻿using System;

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

        public static void OnAssertFailed(string methodName, string message)
        {
            var handler = OnAssertCompleted;
            if (handler != null)
            {
                handler(null, new AssertCompletedEventArgs(TestOutcome.Failed,
                                methodName + " assertion failed." + (string.IsNullOrEmpty(message) ? string.Empty : " " + message)));
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
