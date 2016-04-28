using System.Runtime.InteropServices;

namespace Rubberduck.UnitTesting
{
	
    [ComVisible(true)]
    [ComDefaultInterface(typeof(IAssert))]
    [Guid(ClassId)]
    [ProgId(ProgId)]
    public class PermissiveAssertClass : AssertClass
    {
        private const string ClassId = "40F71F29-D63F-4481-8A7D-E04A4B054501";
        private const string ProgId = "Rubberduck.PermissiveAssertClass";

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
        public override void AreEqual(object expected, object actual, string message = null)
        {
            // vbNullString is marshalled as null. assume value semantics:
            expected = expected ?? string.Empty;
            actual = actual ?? string.Empty;

            if (expected.GetType() != actual.GetType())
            {
                if (!RunTypePromotions(ref expected, ref actual))
                {
                    AssertHandler.OnAssertFailed("AreEqual", message);
                }
            }

            if (expected.Equals(actual))
            {
                AssertHandler.OnAssertSucceeded();
            }
        }

        public override void AreNotEqual(object expected, object actual, string message = null)
        {
            // vbNullString is marshalled as null. assume value semantics:
            expected = expected ?? string.Empty;
            actual = actual ?? string.Empty;

            if (expected.GetType() != actual.GetType())
            {
                if (!RunTypePromotions(ref expected, ref actual))
                {
                    AssertHandler.OnAssertFailed("AreNotEqual", message);
                }
            }

            if (!expected.Equals(actual))
            {
                AssertHandler.OnAssertSucceeded();
            }
        }

        /// <summary>
        /// Runs applicable type promotions for number types.
        /// </summary>
        /// <returns><c>true</c>, if any type promotion was run, <c>false</c> otherwise.</returns>
        /// <param name="expected">the expected value given to the test-method</param>
        /// <param name="actual">the actual value given to the test method</param>
        static bool RunTypePromotions(ref object expected, ref object actual)
        {
            // try promoting integral types first.
            if (expected is ulong && actual is ulong)
            {
                expected = (ulong)expected;
                actual = (ulong)actual;
            }
            // then try promoting to floating point
            else if (expected is double && actual is double)
            {
                expected = (double)expected;
                actual = (double)actual;
            }
            // that shouldn't actually happen, since decimal is the only numeric ValueType in it's category
            // this means we should've gotten the same types earlier in the Assert method
            else if (expected is decimal && actual is decimal)
            {
                expected = (decimal)expected;
                actual = (decimal)actual;
            }
            // worst case scenario for numbers
            // since we're inside VBA though, double is the more appropriate type to compare, 
            // because that is what's used internally anyways, see https://support.microsoft.com/en-us/kb/78113
            else if ((expected is decimal && actual is double) || (expected is double && actual is decimal))
            {
                expected = (double)expected;
                actual = (double)actual;
            }
            // no number-type promotions are applicable.
            else
            {
                // last staw: string "promotion"
                if (expected is string || actual is string)
                {
                    expected = expected.ToString();
                    actual = actual.ToString();
                    return true;
                }
                return false;
            }
            return true;
        }
    }
}

