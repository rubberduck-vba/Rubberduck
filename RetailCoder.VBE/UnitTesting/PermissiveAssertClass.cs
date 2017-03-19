using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Rubberduck.UnitTesting
{	
    [ComVisible(true)]
    [ComDefaultInterface(typeof(IAssert))]
    [Guid(RubberduckGuid.PermissiveAssertClassGuid)]
    [ProgId(RubberduckProgId.PermissiveAssertClassProgId)]
    public class PermissiveAssertClass : AssertClass
    {       
        private static readonly IEqualityComparer<object> PermissiveComparer = new PermissiveObjectComparer();

        /// <summary>
        /// Verifies that two specified objects are equal as considered equal under the loose terms of VBA equality.
        /// As such the assertion fails, if the objects are not equal, even after applying VBA Type promotions.
        /// </summary>
        /// <param name="expected">The expected value.</param>
        /// <param name="actual">The actual value.</param>
        /// <param name="message">An optional message to display if the assertion fails.</param>
        /// <remarks>
        /// contrary to <see cref="AssertClass.AreEqual"/> <paramref name="expected"/> and <paramref name="actual"/> are not required to be of the same type
        /// </remarks>
        public override void AreEqual(object expected, object actual, string message = "")
        {
            // vbNullString is marshalled as null. assume value semantics:
            expected = expected ?? string.Empty;
            actual = actual ?? string.Empty;

            if (!PermissiveComparer.Equals(expected, actual))
            {
                AssertHandler.OnAssertFailed(message);
            }
            AssertHandler.OnAssertSucceeded();
        }

        public override void AreNotEqual(object expected, object actual, string message = "")
        {
            // vbNullString is marshalled as null. assume value semantics:
            expected = expected ?? string.Empty;
            actual = actual ?? string.Empty;

            if (PermissiveComparer.Equals(expected, actual))
            {
                AssertHandler.OnAssertFailed(message);
            }
            AssertHandler.OnAssertSucceeded();
        }

        public override void SequenceEquals(object expected, object actual, string message = "")
        {
            if (!SequenceEquityParametersAreArrays(expected, actual, true))
            {
                return;
            }
            TestArraySequenceEquity((Array)expected, (Array)actual, message, true, PermissiveComparer);
        }

        public override void NotSequenceEquals(object expected, object actual, string message = "")
        {
            if (!SequenceEquityParametersAreArrays(expected, actual, false))
            {
                return;
            }
            TestArraySequenceEquity((Array)expected, (Array)actual, message, false, PermissiveComparer);
        }
    }
}

