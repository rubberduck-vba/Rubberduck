using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Grammar;
using Rubberduck.UI;

namespace Rubberduck.UnitTesting
{
    [ComVisible(true)]
    [ComDefaultInterface(typeof(IAssert))]
    [Guid(RubberduckGuid.AssertClassGuid)]
    [ProgId(RubberduckProgId.AssertClassProgId)]
    public class AssertClass : IAssert
    {
        private static readonly IEqualityComparer<object> DefaultComparer = EqualityComparer<object>.Default;

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
                AssertHandler.OnAssertFailed(message);
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
                AssertHandler.OnAssertFailed(message);
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
            AssertHandler.OnAssertFailed(message);
        }

        public void Succeed()
        {
            AssertHandler.OnAssertSucceeded();
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
                AssertHandler.OnAssertFailed(message);
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
                AssertHandler.OnAssertFailed(message);
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
                AssertHandler.OnAssertFailed(string.Format(RubberduckUI.Assert_ParameterResultFormat, expected, actual, message).Trim());
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
                AssertHandler.OnAssertFailed(string.Format(RubberduckUI.Assert_ParameterResultFormat, expected, actual, message).Trim());
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
                AssertHandler.OnAssertFailed(string.Format(RubberduckUI.Assert_ParameterResultFormat, Tokens.Nothing, actual.GetHashCode(), message).Trim());
                return;
            }
            if (actual == null && expected != null)
            {
                AssertHandler.OnAssertFailed(string.Format(RubberduckUI.Assert_ParameterResultFormat, expected.GetHashCode(), Tokens.Nothing, message).Trim());
                return;
            }
            
            if (ReferenceEquals(expected, actual))
            {
                AssertHandler.OnAssertSucceeded();
            }
            else
            {
                AssertHandler.OnAssertFailed(string.Format(RubberduckUI.Assert_ParameterResultFormat, expected.GetHashCode(), actual.GetHashCode(), message).Trim());
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
                AssertHandler.OnAssertFailed(string.Format(RubberduckUI.Assert_ParameterResultFormat, Tokens.Nothing, Tokens.Nothing, message).Trim());
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

            AssertHandler.OnAssertFailed(string.Format(RubberduckUI.Assert_ParameterResultFormat, expected.GetHashCode(), actual.GetHashCode(), message).Trim());
        }

        public virtual void SequenceEquals(object expected, object actual, string message = "")
        {
            if (!SequenceEquityParametersAreArrays(expected, actual, true))
            {
                return;
            }
            TestArraySequenceEquity((Array)expected, (Array)actual, message, true, DefaultComparer);
        }

        public virtual void NotSequenceEquals(object expected, object actual, string message = "")
        {
            if (!SequenceEquityParametersAreArrays(expected, actual, false))
            {
                return;
            }
            TestArraySequenceEquity((Array)expected, (Array)actual, message, false, DefaultComparer);
        }


        [SuppressMessage("ReSharper", "ExplicitCallerInfoArgument")]
        protected void TestArraySequenceEquity(Array expected, Array actual, string message, bool equals, IEqualityComparer<object> comparer, [CallerMemberName] string methodName = "")
        {
            if (expected.Rank != actual.Rank)
            {
                if (equals)
                {
                    AssertHandler.OnAssertFailed(string.Format(RubberduckUI.Assert_DimensionMismatchFormat, expected.Rank,
                            actual.Rank, message).Trim(), methodName);                    
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
                        AssertHandler.OnAssertFailed(string.Format(RubberduckUI.Assert_LBoundMismatchFormat, rank + 1, expectedBound,
                                actualBound, message).Trim(), methodName);
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
                        AssertHandler.OnAssertFailed(string.Format(RubberduckUI.Assert_UBoundMismatchFormat, rank + 1, expectedBound,
                                actualBound, message).Trim(), methodName);
                        return;
                    }
                    AssertHandler.OnAssertSucceeded();
                }
            }

            var flattenedExpected = expected.Cast<object>().ToList();
            var flattenedActual = actual.Cast<object>().ToList();
            if (!flattenedActual.SequenceEqual(flattenedExpected, comparer))
            {
                if (equals)
                {
                    AssertHandler.OnAssertFailed(message, methodName);
                }
                AssertHandler.OnAssertSucceeded();
            }

            if (!equals)
            {
                AssertHandler.OnAssertFailed(message, methodName);
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
                    string.Format(RubberduckUI.Assert_UnexpectedArrayFormat, equals ? "Assert.SequenceEquals" : "Assert.NotSequenceEquals"));
                return false;
            }

            if (!ReferenceOrValueTypesMatch(expectedType, actualType))
            {
                return false;
            }

            if (expectedType.IsCOMObject && actualType.IsCOMObject)
            {
                AssertHandler.OnAssertInconclusive(
                    string.Format(RubberduckUI.Assert_UnexpectedReferenceComparisonFormat, equals ? "Assert.AreSame" : "Assert.AreNotSame"));
                return false;
            }

            if (expectedType != actualType)
            {
                AssertHandler.OnAssertInconclusive(RubberduckUI.Assert_MismatchedTypes);
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
                    string.Format(RubberduckUI.Assert_UnexpectedValueComparisonFormat, same ? "Assert.AreEqual" : "Assert.AreNotEqual"));
                return false;
            }
            return true;
        }

        protected bool SequenceEquityParametersAreArrays(object expected, object actual, bool equals)
        {
            var expectedType = expected?.GetType();
            var actualType = actual?.GetType();

            if (expectedType == null && actualType == null)
            {
                AssertHandler.OnAssertInconclusive(
                    string.Format(RubberduckUI.Assert_UnexpectedNullArraysFormat, equals ? "Assert.AreSame" : "Assert.AreNotSame"));
                return false;
            }

            if (!ReferenceOrValueTypesMatch(expectedType, actualType))
            {
                return false;
            }

            if (expectedType != null && !expectedType.IsArray && actualType != null && actualType.IsArray)
            {
                AssertHandler.OnAssertInconclusive(string.Format(RubberduckUI.Assert_ParameterIsNotArrayFormat, "[Expected]"));
                return false;
            }

            if (actualType != null && !actualType.IsArray && expectedType != null && expectedType.IsArray)
            {
                AssertHandler.OnAssertInconclusive(string.Format(RubberduckUI.Assert_ParameterIsNotArrayFormat, "[Actual]"));
                return false;
            }

            if (actualType != null && !actualType.IsArray && (expectedType == null || expectedType.IsArray))
            {
                AssertHandler.OnAssertInconclusive(RubberduckUI.Assert_NeitherParameterIsArray);
                return false;
            }

            return true;
        }

        private bool ReferenceOrValueTypesMatch(Type expectedType, Type actualType)
        {
            if (expectedType != null && !expectedType.IsCOMObject && (actualType == null || actualType.IsCOMObject))
            {
                AssertHandler.OnAssertInconclusive(RubberduckUI.Assert_ValueReferenceMismatch);
                return false;
            }

            if (actualType != null && !actualType.IsCOMObject && (expectedType == null || expectedType.IsCOMObject))
            {
                AssertHandler.OnAssertInconclusive(RubberduckUI.Assert_ReferenceValueMismatch);
                return false;
            }
            return true;
        }
    }
}
