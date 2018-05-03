using System.ComponentModel;
using System.Runtime.InteropServices;

// The parameters on RD's public interfaces are following VBA conventions not C# conventions to stop the
// obnoxious "Can I haz all identifiers with the same casing" behavior of the VBE.
// ReSharper disable InconsistentNaming

namespace Rubberduck.UnitTesting
{
    [ComVisible(true)]
    [Guid(RubberduckGuid.IVerifyGuid)]
    [EditorBrowsable(EditorBrowsableState.Always)]
    public interface IVerify
    {
        /// <summary>
        /// Verifies that the faked procedure was called a minimum number of times.
        /// </summary>
        /// <param name="Invocations">Expected minimum invocations.</param>
        /// <param name="Message">An optional message to display if the verification fails.</param>
        [DispId(1)]
        [Description("Verifies that the faked procedure was called a minimum number of times.")]
        void AtLeast(int Invocations, string Message = "");

        /// <summary>
        /// Verifies that the faked procedure was called one or more times.
        /// </summary>
        /// <param name="Message">An optional message to display if the verification fails.</param>
        [DispId(2)]
        [Description("Verifies that the faked procedure was called one or more times.")]
        void AtLeastOnce(string Message = "");

        /// <summary>
        /// Verifies that the faked procedure was called a maximum number of times.
        /// </summary>
        /// <param name="Invocations">Expected maximum invocations.</param>
        /// <param name="Message">An optional message to display if the verification fails.</param>
        [DispId(3)]
        [Description("Verifies that the faked procedure was called a maximum number of times.")]
        void AtMost(int Invocations, string Message = "");

        /// <summary>
        /// Verifies that the faked procedure was not called or was only called once.
        /// </summary>
        /// <param name="Message">An optional message to display if the verification fails.</param>
        [DispId(4)]
        [Description("Verifies that the faked procedure was not called or was only called once.")]
        void AtMostOnce(string Message = "");

        /// <summary>
        /// Verifies that number of times the faked procedure was called falls within the supplied range.
        /// </summary>
        /// <param name="Minimum">Expected minimum invocations.</param>
        /// <param name="Maximum">Expected maximum invocations.</param>
        /// <param name="Message">An optional message to display if the verification fails.</param>
        [DispId(5)]
        [Description("Verifies that number of times the faked procedure was called falls within the supplied range.")]
        void Between(int Minimum, int Maximum, string Message = "");

        /// <summary>
        /// Verifies that the faked procedure was called a specific number of times.
        /// </summary>
        /// <param name="Invocations">Expected invocations.</param>
        /// <param name="Message">An optional message to display if the verification fails.</param>
        [DispId(6)]
        [Description("Verifies that the faked procedure was called a specific number of times.")]
        void Exactly(int Invocations, string Message = "");

        /// <summary>
        /// Verifies that the faked procedure is never called.
        /// </summary>
        /// <param name="Message">An optional message to display if the verification fails.</param>
        [DispId(7)]
        [Description("Verifies that the faked procedure is never called.")]
        void Never(string Message = "");

        /// <summary>
        /// Verifies that the faked procedure is called exactly one time.
        /// </summary>
        /// <param name="Message">An optional message to display if the verification fails.</param>
        [DispId(8)]
        [Description("Verifies that the faked procedure is called exactly one time.")]
        void Once(string Message = "");

        /// <summary>
        /// Verifies that a given parameter to the faked procedure matches a specific value.
        /// </summary>
        /// <param name="Parameter">The name of the parameter to verify. Case insensitive.</param>
        /// <param name="Value">The expected value of the parameter.</param>
        /// <param name="Invocation">The invocation to test against. Optional - defaults to the first invocation.</param>
        /// <param name="Message">An optional message to display if the verification fails.</param>
        [DispId(9)]
        [Description("Verifies that a given parameter to the faked procedure matches a specific value.")]
        void Parameter(string Parameter, object Value, int Invocation = 1, string Message = "");

        /// <summary>
        /// Verifies that the value of a given parameter to the faked procedure falls within a specified range.
        /// </summary>
        /// <param name="Parameter">The name of the parameter to verify. Case insensitive.</param>
        /// <param name="Minimum">Expected minimum value.</param>
        /// <param name="Maximum">Expected maximum value.</param>
        /// <param name="Invocation">The invocation to test against. Optional - defaults to the first invocation.</param>
        /// <param name="Message">An optional message to display if the verification fails.</param>
        [DispId(10)]
        [Description("Verifies that the value of a given parameter to the faked procedure falls within a specified range.")]
        void ParameterInRange(string Parameter, double Minimum, double Maximum, int Invocation = 1, string Message = "");

        /// <summary>
        /// Verifies that an optional parameter was passed to the faked procedure. The value is not evaluated.
        /// </summary>
        /// <param name="Parameter">The name of the parameter to verify. Case insensitive.</param>
        /// <param name="Invocation">The invocation to test against. Optional - defaults to the first invocation.</param>
        /// <param name="Message">An optional message to display if the verification fails.</param>
        [DispId(11)]
        [Description("Verifies that an optional parameter was passed to the faked procedure. The value is not evaluated.")]
        void ParameterIsPassed(string Parameter, int Invocation = 1, string Message = "");

        /// <summary>
        /// Verifies that the passed value of a given parameter is of a given type.
        /// </summary>
        /// <param name="Parameter">The name of the parameter to verify. Case insensitive.</param>
        /// <param name="TypeName">The expected type as it would be returned by VBA.TypeName. Case insensitive.</param>
        /// <param name="Invocation">The invocation to test against. Optional - defaults to the first invocation.</param>
        /// <param name="Message">An optional message to display if the verification fails.</param>
        [DispId(12)]
        [Description("Verifies that the passed value of a given parameter is of a given type.")]
        void ParameterIsType(string Parameter, string TypeName, int Invocation = 1, string Message = "");
    }
}
