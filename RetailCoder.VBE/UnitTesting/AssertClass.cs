using System;
using System.Linq;
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
        public void IsTrue(bool condition, string message = "")
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
        public void IsFalse(bool condition, string message = "")
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
        public void Inconclusive(string message = "")
        {
            AssertHandler.OnAssertInconclusive(message);
        }

        /// <summary>
        /// Fails the assertion without checking any conditions.
        /// </summary>
        /// <param name="message">An optional message to display.</param>
        public void Fail(string message = "")
        {
            AssertHandler.OnAssertFailed("Fail", message);
        }

        /// <summary>
        /// Verifies that the specified object is <c>Nothing</c>. The assertion fails if it is not <c>Nothing</c>.
        /// </summary>
        /// <param name="value">The object to verify.</param>
        /// <param name="message">An optional message to display if the assertion fails.</param>
        public void IsNothing(object value, string message = "")
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
        public void IsNotNothing(object value, string message = "")
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
        public virtual void AreEqual(object expected, object actual, string message = "")
        {
            // vbNullString is marshaled as a null. assume value semantics:
            expected = expected ?? string.Empty;
            actual = actual ?? string.Empty;

            if (!ValueEquityAssertTypesMatch(expected, actual, true))
            {
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
        public virtual void AreNotEqual(object expected, object actual, string message = "")
        {
            // vbNullString is marshaled as a null. assume value semantics:
            expected = expected ?? string.Empty;
            actual = actual ?? string.Empty;

            if (!ValueEquityAssertTypesMatch(expected, actual, false))
            {
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
        public void AreSame(object expected, object actual, string message = "")
        {
            if (!ReferenceEquityAssertTypesMatch(expected, actual, true))
            {
                return;
            }

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
        public void AreNotSame(object expected, object actual, string message = "")
        {
            if (!ReferenceEquityAssertTypesMatch(expected, actual, false))
            {
                return;
            }

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

        public void SequenceEquals(object expected, object actual, string message = "")
        {
            if (!SequenceEquityParametersAreArrays(expected, actual, true))
            {
                return;
            }
            TestArraySequenceEquity((Array)expected, (Array)actual, message, true);
        }

        public void NotSequenceEquals(object expected, object actual, string message = "")
        {
            if (!SequenceEquityParametersAreArrays(expected, actual, false))
            {
                return;
            }
            TestArraySequenceEquity((Array)expected, (Array)actual, message, false);
        }

        private void TestArraySequenceEquity(Array expected, Array actual, string message, bool equals)
        {
            if (expected.Rank != actual.Rank)
            {
                if (equals)
                {
                    AssertHandler.OnAssertFailed("SequenceEquals",
                        string.Format("expected has {0} dimensions; actual has {1} dimensions. {2} ", expected.Rank,
                            actual.Rank, message).Trim());
                    return;
                }
                AssertHandler.OnAssertSucceeded();
            }

            for (var rank = 0; rank < expected.Rank; rank++)
            {
                var expectedBound = expected.GetLowerBound(rank);
                var actualBound = actual.GetLowerBound(rank);
                if (expectedBound != actualBound)
                {
                    if (equals)
                    {
                        AssertHandler.OnAssertFailed("SequenceEquals",
                            string.Format("Dimension {0}: expected has an LBound of {1}; actual has an LBound of {2}. {3} ", rank + 1, expectedBound,
                                actualBound, message).Trim());
                        return;
                    }
                    AssertHandler.OnAssertSucceeded();
                }

                expectedBound = expected.GetUpperBound(rank);
                actualBound = actual.GetUpperBound(rank);
                if (expectedBound != actualBound)
                {
                    if (equals)
                    {
                        AssertHandler.OnAssertFailed("SequenceEquals",
                            string.Format("Dimension {0}: expected has a UBound of {1}; actual has a UBound of {2}. {3} ", rank + 1, expectedBound,
                                actualBound, message).Trim());
                        return;
                    }
                    AssertHandler.OnAssertSucceeded();
                }
            }

            var flattenedExpected = expected.Cast<object>().ToList();
            var flattenedActual = actual.Cast<object>().ToList();
            if (!flattenedActual.SequenceEqual(flattenedExpected))
            {
                if (equals)
                {
                    AssertHandler.OnAssertFailed("SequenceEquals", message);
                }
                AssertHandler.OnAssertSucceeded();
            }

            if (!equals)
            {
                AssertHandler.OnAssertFailed("NotSequenceEquals", message);
            }
            AssertHandler.OnAssertSucceeded();
        }

        private bool ValueEquityAssertTypesMatch(object expected, object actual, bool equals)
        {
            var expectedType = expected.GetType();
            var actualType = actual.GetType();

            if (expectedType.IsArray && actualType.IsArray)
            {
                AssertHandler.OnAssertInconclusive(
                    string.Format("[expected] and [actual] are arrays. Consider using {0}.",
                        equals ? "Assert.SequenceEquals" : "Assert.NotSequenceEquals"));
                return false;
            }

            if (!ReferenceOrValueTypesMatch(expectedType, actualType))
            {
                return false;
            }

            if (expectedType.IsCOMObject && actualType.IsCOMObject)
            {
                AssertHandler.OnAssertInconclusive(
                    string.Format("[expected] and [actual] are reference types. Consider using {0}.",
                        equals ? "Assert.AreSame" : "Assert.AreNotSame"));
                return false;
            }

            if (expectedType != actualType)
            {
                AssertHandler.OnAssertInconclusive("[expected] and [actual] values are not the same type.");
                return false;
            }
            return true;
        }

        private bool ReferenceEquityAssertTypesMatch(object expected, object actual, bool same)
        {
            var expectedType = expected?.GetType();
            var actualType = actual?.GetType();

            if ((expectedType == null && actualType == null) || 
                ((expectedType == null || expectedType.IsCOMObject) && (actualType == null || actualType.IsCOMObject)))
            {
                return true;
            }

            if (!ReferenceOrValueTypesMatch(expectedType, actualType))
            {
                return false;
            }

            if (expectedType != null && !expectedType.IsCOMObject && actualType != null && !actualType.IsCOMObject)
            {
                AssertHandler.OnAssertInconclusive(
                    string.Format("[expected] and [actual] are value types. Consider using {0}.",
                        same ? "Assert.AreEqual" : "Assert.AreNotEqual"));
                return false;
            }
            return true;
        }

        private bool SequenceEquityParametersAreArrays(object expected, object actual, bool equals)
        {
            var expectedType = expected?.GetType();
            var actualType = actual?.GetType();

            if (expectedType == null && actualType == null)
            {
                AssertHandler.OnAssertInconclusive(
                    string.Format("[expected] and [actual] are Nothing. Consider using {0}.",
                        equals ? "Assert.AreSame" : "Assert.AreNotSame"));
                return false;
            }

            if (!ReferenceOrValueTypesMatch(expectedType, actualType))
            {
                return false;
            }

            if (expectedType != null && !expectedType.IsArray && actualType != null && actualType.IsArray)
            {
                AssertHandler.OnAssertInconclusive("[expected] is a not an array.");
                return false;
            }

            if (actualType != null && !actualType.IsArray && expectedType != null && expectedType.IsArray)
            {
                AssertHandler.OnAssertInconclusive("[actual] is a not an array.");
                return false;
            }

            if (actualType != null && !actualType.IsArray && (expectedType == null || expectedType.IsArray))
            {
                AssertHandler.OnAssertInconclusive("Neither [expected] and [actual] is an array.");
                return false;
            }

            return true;
        }

        private bool ReferenceOrValueTypesMatch(Type expectedType, Type actualType)
        {
            if (expectedType != null && !expectedType.IsCOMObject && (actualType == null || actualType.IsCOMObject))
            {
                AssertHandler.OnAssertInconclusive("[expected] is a value type and [actual] is a reference type.");
                return false;
            }

            if (actualType != null && !actualType.IsCOMObject && (expectedType == null || expectedType.IsCOMObject))
            {
                AssertHandler.OnAssertInconclusive("[expected] is a reference type and [actual] is a value type.");
                return false;
            }
            return true;
        }
    }
}
