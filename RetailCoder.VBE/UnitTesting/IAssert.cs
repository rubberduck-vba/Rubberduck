using System.Runtime.InteropServices;

// The parameters on RD's public interfaces are following VBA conventions not C# conventions to stop the
// obnoxious "Can I haz all identifiers with the same casing" behavior of the VBE.
// ReSharper disable InconsistentNaming

namespace Rubberduck.UnitTesting
{
    [ComVisible(true)]
    [Guid("56EC4DBF-508F-46A5-8067-4F4354F4632C")]
    public interface IAssert
    {
        /// <summary>
        /// Verifies that the specified condition is <c>true</c>. The assertion fails if the condition is <c>false</c>.
        /// </summary>
        /// <param name="Condition">Any Boolean value or expression.</param>
        /// <param name="Message">An optional message to display if the assertion fails.</param>
        void IsTrue(bool Condition, string Message = "");

        /// <summary>
        /// Verifies that the specified condition is <c>false</c>. The assertion fails if the condition is <c>true</c>.
        /// </summary>
        /// <param name="Condition">Any Boolean value or expression.</param>
        /// <param name="Message">An optional message to display if the assertion fails.</param>
        void IsFalse(bool Condition, string Message = "");

        /// <summary>
        /// Indicates that the assertion cannot be verified.
        /// </summary>
        /// <param name="Message">An optional message to display.</param>
        void Inconclusive(string Message = "");

        /// <summary>
        /// Fails the assertion without checking any conditions.
        /// </summary>
        /// <param name="Message">An optional message to display.</param>
        void Fail(string Message = "");

        /// <summary>
        /// Verifies that the specified object is <c>Nothing</c>. The assertion fails if it is not <c>Nothing</c>.
        /// </summary>
        /// <param name="Value">The object to verify.</param>
        /// <param name="Message">An optional message to display if the assertion fails.</param>
        void IsNothing(object Value, string Message = "");

        /// <summary>
        /// Verifies that the specified object is not <c>Nothing</c>. The assertion fails if it is <c>Nothing</c>.
        /// </summary>
        /// <param name="Value">The object to verify.</param>
        /// <param name="Message">An optional message to display if the assertion fails.</param>
        void IsNotNothing(object Value, string Message = "");

        /// <summary>
        /// Verifies that two specified objects are equal. The assertion fails if the objects are not equal.
        /// </summary>
        /// <param name="Expected">The expected value.</param>
        /// <param name="Actual">The actual value.</param>
        /// <param name="Message">An optional message to display if the assertion fails.</param>
        /// <remarks>
        /// <paramref name="Expected"/> and <paramref name="Actual"/> must be the same type.
        /// </remarks>
        void AreEqual(object Expected, object Actual, string Message = "");

        /// <summary>
        /// Verifies that two specified objects are not equal. The assertion fails if the objects are equal.
        /// </summary>
        /// <param name="Expected">The expected value.</param>
        /// <param name="Actual">The actual value.</param>
        /// <param name="Message">An optional message to display if the assertion fails.</param>
        /// <remarks>
        /// <paramref name="Expected"/> and <paramref name="Actual"/> must be the same type.
        /// </remarks>
        void AreNotEqual(object Expected, object Actual, string Message = "");

        /// <summary>
        /// Verifies that two specified object variables refer to the same object. The assertion fails if they refer to different objects.
        /// </summary>
        /// <param name="Expected">The expected reference.</param>
        /// <param name="Actual">The actual reference.</param>
        /// <param name="Message">An optional message to display if the assertion fails.</param>
        void AreSame(object Expected, object Actual, string Message = "");

        /// <summary>
        /// Verifies that two specified object variables refer to different objects. The assertion fails if they refer to the same object.
        /// </summary>
        /// <param name="Expected">The expected reference.</param>
        /// <param name="Actual">The actual reference.</param>
        /// <param name="Message">An optional message to display if the assertion fails.</param>
        void AreNotSame(object Expected, object Actual, string Message = "");


        void SequenceEquals(object Expected, object Actual, string Message = "");
        void NotSequenceEquals(object Expected, object Actual, string Message = "");
    }
}
