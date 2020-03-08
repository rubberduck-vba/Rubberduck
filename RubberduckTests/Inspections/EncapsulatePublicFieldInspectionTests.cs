using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    [Category("EncapsulatePublicFieldInspection")]
    public class EncapsulatePublicFieldInspectionTests : InspectionTestsBase
    {
        [Test]
        public void PublicField_ReturnsResult()
        {
            const string inputCode =
                @"Public fizz As Boolean";
            
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        public void GlobalField_ReturnsResult()
        {
            const string inputCode =
                @"Global fizz As Boolean";
            
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        public void MultiplePublicFields_ReturnMultipleResult()
        {
            const string inputCode =
                @"Public fizz As Boolean
Public buzz As Integer, _
       bazz As Integer";

            Assert.AreEqual(3, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        public void PrivateField_DoesNotReturnResult()
        {
            const string inputCode =
                @"Private fizz As Boolean";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        public void PublicNonField_DoesNotReturnResult()
        {
            const string inputCode =
                @"Public Sub Foo(ByRef arg1 As String)
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        public void PublicField_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"'@Ignore EncapsulatePublicField
Public fizz As Boolean";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        public void GlobalField_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"'@Ignore EncapsulatePublicField
Global fizz As Boolean";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        public void InspectionName()
        {
            var inspection = new EncapsulatePublicFieldInspection(null);

            Assert.AreEqual(nameof(EncapsulatePublicFieldInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new EncapsulatePublicFieldInspection(state);
        }
    }
}
