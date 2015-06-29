using System.ComponentModel;
using System.Runtime.InteropServices;

namespace Rubberduck.UnitTesting
{
    [ComVisible(true)]
    [ComDefaultInterface(typeof(IAssert))]
    [Guid(ClassId)]
    [ProgId(ProgId)]
    public class AssertClass : IAssert
    {
        private const string ClassId = "69E194DA-43F0-3B33-B105-9B8188A6F040";
        private const string ProgId = "Rubberduck.AssertClass";

        /// <summary>
        /// Verifies that the specified condition is <c>true</c>. The assertion fails if the condition is <c>false</c>.
        /// </summary>
        /// <param name="condition">Any Boolean value or expression.</param>
        /// <param name="message">An optional message to display if the assertion fails.</param>
        public void IsTrue(bool condition, string message = null)
        {
            if (condition)
            {
                AssertHandler.OnAssertSucceeded();
            }
            else
            {
                AssertHandler.OnAssertFailed("IsTrue", message);
            }
        }

        /// <summary>
        /// Verifies that the specified condition is <c>false</c>. The assertion fails if the condition is <c>true</c>.
        /// </summary>
        /// <param name="condition">Any Boolean value or expression.</param>
        /// <param name="message">An optional message to display if the assertion fails.</param>
        public void IsFalse(bool condition, string message = null)
        {
            if (!condition)
            {
                AssertHandler.OnAssertSucceeded();
            }
            else
            {
                AssertHandler.OnAssertFailed("IsFalse", message);
            }
        }

        /// <summary>
        /// Indicates that the assertion cannot be verified.
        /// </summary>
        /// <param name="message">An optional message to display.</param>
        public void Inconclusive(string message = null)
        {
            AssertHandler.OnAssertInconclusive(message);
        }

        /// <summary>
        /// Fails the assertion without checking any conditions.
        /// </summary>
        /// <param name="message">An optional message to display.</param>
        public void Fail(string message = null)
        {
            AssertHandler.OnAssertFailed("Fail", message);
        }

        /// <summary>
        /// Verifies that the specified object is <c>Nothing</c>. The assertion fails if it is not <c>Nothing</c>.
        /// </summary>
        /// <param name="value">The object to verify.</param>
        /// <param name="message">An optional message to display if the assertion fails.</param>
        public void IsNothing(object value, string message = null)
        {
            if (value == null)
            {
                AssertHandler.OnAssertSucceeded();
            }
            else
            {
                AssertHandler.OnAssertFailed("IsNothing", message);
            }
        }

        /// <summary>
        /// Verifies that the specified object is not <c>Nothing</c>. The assertion fails if it is <c>Nothing</c>.
        /// </summary>
        /// <param name="value">The object to verify.</param>
        /// <param name="message">An optional message to display if the assertion fails.</param>
        public void IsNotNothing(object value, string message = null)
        {
            if (value != null)
            {
                AssertHandler.OnAssertSucceeded();
            }
            else
            {
                AssertHandler.OnAssertFailed("IsNotNothing", message);
            }
        }

        /// <summary>
        /// Verifies that two specified objects are equal. The assertion fails if the objects are not equal.
        /// </summary>
        /// <param name="expected">The expected value.</param>
        /// <param name="actual">The actual value.</param>
        /// <param name="message">An optional message to display if the assertion fails.</param>
        /// <remarks>
        /// <paramref name="expected"/> and <paramref name="actual"/> must be the same type.
        /// </remarks>
        public void AreEqual(object expected, object actual, string message = null)
        {
            // vbNullString is marshaled as a null. assume value semantics:
            expected = expected ?? string.Empty;
            actual = actual ?? string.Empty;

            if (expected.GetType() != actual.GetType())
            {
                AssertHandler.OnAssertInconclusive("[expected] and [actual] values are not the same type.");
                return;
            }

            if (expected.Equals(actual))
            {
                AssertHandler.OnAssertSucceeded();
            }
            else
            {
                AssertHandler.OnAssertFailed("AreEqual", string.Concat("expected: ", expected.ToString(), "; actual: ", actual.ToString(), ". ", message));
            }
        }

        /// <summary>
        /// Verifies that two specified objects are not equal. The assertion fails if the objects are equal.
        /// </summary>
        /// <param name="expected">The expected value.</param>
        /// <param name="actual">The actual value.</param>
        /// <param name="message">An optional message to display if the assertion fails.</param>
        /// <remarks>
        /// <paramref name="expected"/> and <paramref name="actual"/> must be the same type.
        /// </remarks>
        public void AreNotEqual(object expected, object actual, string message = null)
        {
            // vbNullString is marshaled as a null. assume value semantics:
            expected = expected ?? string.Empty;
            actual = actual ?? string.Empty;

            if (expected.GetType() != actual.GetType())
            {
                AssertHandler.OnAssertInconclusive("[expected] and [actual] values are not the same type.");
                return;
            }

            if (!expected.Equals(actual))
            {
                AssertHandler.OnAssertSucceeded();
            }
            else
            {
                AssertHandler.OnAssertFailed("AreNotEqual", string.Concat("expected: ", expected.ToString(), "; actual: ", actual.ToString(), ". ", message));
            }
        }

        /// <summary>
        /// Verifies that two specified object variables refer to the same object. The assertion fails if they refer to different objects.
        /// </summary>
        /// <param name="expected">The expected reference.</param>
        /// <param name="actual">The actual reference.</param>
        /// <param name="message">An optional message to display if the assertion fails.</param>
        public void AreSame(object expected, object actual, string message = null)
        {
            if (expected == null && actual != null)
            {
                AssertHandler.OnAssertFailed("AreSame", string.Concat("expected: Nothing; actual: ", actual.GetHashCode(), ". ", message));
                return;
            }
            if (actual == null && expected != null)
            {
                AssertHandler.OnAssertFailed("AreSame", string.Concat("expected: ", expected.GetHashCode(), "; actual: Nothing. ", message));
                return;
            }
            
            if (ReferenceEquals(expected, actual))
            {
                AssertHandler.OnAssertSucceeded();
            }
            else
            {
                AssertHandler.OnAssertFailed("AreSame", string.Concat("expected: ", expected.GetHashCode(), "; actual: ", actual.GetHashCode(), ". ", message));
            }
        }

        /// <summary>
        /// Verifies that two specified object variables refer to different objects. The assertion fails if they refer to the same object.
        /// </summary>
        /// <param name="expected">The expected reference.</param>
        /// <param name="actual">The actual reference.</param>
        /// <param name="message">An optional message to display if the assertion fails.</param>
        public void AreNotSame(object expected, object actual, string message = null)
        {
            if (expected == null && actual == null)
            {
                AssertHandler.OnAssertFailed("AreNotSame", string.Concat("expected: Nothing; actual: Nothing. ", message));
                return;
            }
            if (expected == null || actual == null)
            {
                AssertHandler.OnAssertSucceeded();
                return;
            }
            
            if (!ReferenceEquals(expected, actual))
            {
                AssertHandler.OnAssertSucceeded();
                return;
            }

            AssertHandler.OnAssertFailed("AreNotSame", string.Concat("expected: ", expected.GetHashCode(), "; actual: ", actual.GetHashCode(), ". ", message));
        }
    }
}
