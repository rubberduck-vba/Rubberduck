using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class HungarianNotationInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void HungarianNotation_ReturnsResult_VariableWithThreeLetterPrefix()
        {
            const string inputCode =
                @"Sub Hungarian()
    Dim strFoo As String
End Sub";

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void HungarianNotation_ReturnsResult_VariableWithOneLetterPrefix()
        {
            const string inputCode =
                @"Sub Hungarian()
    Dim oFoo As Object
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void HungarianNotation_ReturnsResult_ForClass()
        {
            const string inputCode =
                @"Sub Test()
    Debug.Print ""Ez egy objektum""
End Sub";
            Assert.AreEqual(1, InspectionResultsForModules(("clsMagyar", inputCode, ComponentType.ClassModule)).Count());
        }

        [Test]
        [Category("Inspections")]
        public void HungarianNotation_DoesNotReturnsResult_AllLowerCase()
        {
            const string inputCode =
                @"Sub NoHungarianHere()
    Dim strong As Variant
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void HungarianNotation_DoesNotReturnsResult_UpperCaseFirstLetter()
        {
            const string inputCode =
                @"Option Explicit";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void HungarianNotation_DoesNotReturnsResult_ThreeLetterVariable()
        {
            const string inputCode =
                @"Sub InExcelSomewhere()
    Dim col As Long
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void HungarianNotation_DoesNotReturnResult_WhenIgnored()
        {
            const string inputCode =
                @"Sub MagyarRendbenVan()
    '@Ignore HungarianNotation
    Dim strFoo As Variant
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void HungarianNotation_DoesNotReturnResult_WhenWhitelisted()
        {
            const string inputCode =
                @"Sub Feherlista()
    Dim oRange As Object
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void HungarianNotation_DoesNotReturnResult_Ignored()
        {
            const string inputCode =
                @"Sub Hungarian()
'@Ignore HungarianNotation
    Dim oFoo As Object
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void HungarianNotation_DoesNotReturnResult_LibraryFunctionParameters()
        {
            const string inputCode =
                @"
Private Declare Function GetUserName Lib ""advapi32.dll"" Alias ""GetUserNameA"" (ByVal lpBuffer As String, nSize As Long) As Long
";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void HungarianNotation_DoesNotReturnResult_LibraryFunction()
        {
            const string inputCode =
                @"
Private Declare Sub chkVoid Lib ""somelib.dll"" Alias ""chkVoidA"" (number As Long)
";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new HungarianNotationInspection(state, UseMeaningfulNameInspectionTests.GetInspectionSettings().Object);
        }
    }
}
