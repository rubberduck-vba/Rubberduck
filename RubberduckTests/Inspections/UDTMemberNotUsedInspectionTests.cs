using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    class UDTMemberNotUsedInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        [Category(nameof(UDTMemberNotUsedInspection))]
        public void ReturnsZeroResult()
        {
            const string inputCode =
@"
Option Explicit

Private Type TUnitTest
    FirstVal As Long
    SecondVal As Long
End Type

Private this As TUnitTest

Private Sub TestSub(testVal As Long)
    this.FirstVal = testVal * 2
    this.SecondVal = testVal
End Sub
";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        [Category(nameof(UDTMemberNotUsedInspection))]
        public void ReturnsSingleResult()
        {
            const string inputCode =
@"
Option Explicit

Private Type TUnitTest
    FirstVal As Long
    SecondVal As Long
End Type

Private this As TUnitTest

Private Sub TestSub(testVal As Long)
    this.FirstVal = testVal
End Sub
";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        [Category(nameof(UDTMemberNotUsedInspection))]
        public void ReturnsManyResults()
        {
            const string inputCode =
@"
Option Explicit

Private Type TUnitTest
    FirstVal As Long
    SecondVal As Long
    ThirdVal As Long
End Type

Private this As TUnitTest

Private Sub TestSub(testVal As Long)
    this.SecondVal = testVal
End Sub
";
            Assert.AreEqual(2, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        [Category(nameof(UDTMemberNotUsedInspection))]
        public void ReturnsResultForNestedUDTMember()
        {
            const string inputCode =
@"
Option Explicit

Private Type TPair
    IDNumber As Long
    IDName As String
End Type

Private Type TUnitTest
    ID_Name_Pair As TPair
    SecondVal As Long
End Type

Private this As TUnitTest

Private Sub TestSub(testVal As Long)
    this.ID_Name_Pair.IDNumber = testVal
    this.SecondVal = testVal * 2
End Sub
";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        [Category(nameof(UDTMemberNotUsedInspection))]
        public void RespectsIgnoreAnnotation()
        {
            const string inputCode =
@"
Option Explicit

Private Type TUnitTest
    FirstVal As Long
    '@Ignore UDTMemberNotUsed
    SecondVal As Long
End Type

Private this As TUnitTest

Private Sub TestSub(testVal As Long)
    this.FirstVal = testVal
End Sub
";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [TestCase("'@IgnoreModule")]
        [TestCase("'@IgnoreModule UDTMemberNotUsed")]
        [Category("Inspections")]
        [Category(nameof(UDTMemberNotUsedInspection))]
        public void RespectsIgnoreModuleAnnotation(string annotation)
        {
            var inputCode =
$@"
{annotation}
Option Explicit

Private Type TUnitTest
    FirstVal As Long
    SecondVal As Long
End Type

Private this As TUnitTest

Private Sub TestSub(testVal As Long)
    this.FirstVal = testVal
End Sub
";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new UDTMemberNotUsedInspection(state);
        }
    }
}
