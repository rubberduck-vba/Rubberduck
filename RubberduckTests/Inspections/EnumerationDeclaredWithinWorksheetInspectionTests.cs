using System.Linq;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using NUnit.Framework;

namespace RubberduckTests.Inspections
{
    class PublicEnumerationDeclaredWithinWorksheetInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void EnumerationDeclaredWithinWorksheet_InspectionName()
        {
            var inspection = new PublicEnumerationDeclaredWithinWorksheetInspection(null, null);

            Assert.AreEqual(nameof(PublicEnumerationDeclaredWithinWorksheetInspection), inspection.Name);
        }

        [Test]
        [Category("Inspections")]
        [Category(nameof(PublicEnumerationDeclaredWithinWorksheetInspection))]
        [TestCase(ComponentType.Document, 1)]
        [TestCase(ComponentType.StandardModule, 0)]
        [TestCase(ComponentType.ClassModule, 0)]
        [TestCase(ComponentType.UserForm, 0)]
        public void Project_with_public_enumeration_flags_only_enum_declared_within_worksheets(ComponentType componentType, int expected)
        {
            const string code = @"Option Explicit
Public Enum DeclaredEnum
    wsMember1 = 0
    wsMember1 = 1
End Enum
";

            var inspectionResults = InspectionResultsForModules(
                (componentType.ToString(), code, componentType));

            int actual = inspectionResults.Count();

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Inspections")]
        [Category(nameof(PublicEnumerationDeclaredWithinWorksheetInspection))]
        public void Project_with_only_private_worksheet_enumeration_does_not_return_result()
        {
            const string worksheetCode = @"Option Explicit
Private Enum WorksheetEnum
    wsMember1 = 0
    wsMember1 = 1
End Enum
";

            var inspectionResults = InspectionResultsForModules(
                ("FirstSheet", worksheetCode, ComponentType.Document));

            int actual = inspectionResults.Count();

            Assert.AreEqual(0, actual);
        }

        [Test]
        [Category("Inspections")]
        [Category(nameof(PublicEnumerationDeclaredWithinWorksheetInspection))]
        public void Project_with_public_and_private_enumeration_declared_within_worksheets_returns_only_public_declarations()
        {
            const string publicDeclaration = @"Option Explicit
Public Enum WorksheetEnum
    wsMember1 = 0
    wsMember2 = 1
End Enum
";
            const string privateDeclaration = @"Option Explicit
Private Enum WorksheetEnum
    wsMember1 = 0
    wsMember2 = 1
End Enum
";

            var inspectionResults = InspectionResultsForModules(
                ("FirstPublicSheet", publicDeclaration, ComponentType.Document),
                ("FirstPrivateSheet", privateDeclaration, ComponentType.Document),
                ("SecondPrivateSheet", privateDeclaration, ComponentType.Document),
                ("SecondPublicSheet", publicDeclaration, ComponentType.Document),
                ("ThirdPrivateSheet", privateDeclaration, ComponentType.Document));

            int actual = inspectionResults.Count();

            Assert.AreEqual(2, actual);
        }

        [Test]
        [Category("Inspections")]
        [Category(nameof(PublicEnumerationDeclaredWithinWorksheetInspection))]
        public void Private_type_declared_within_worksheet_has_no_inspection_result()
        {
            const string code = @"Option Explicit

Private Type THelper
    Name As String
    Address As String
End Type

Private this as THelper
";

            var inspectionResults = InspectionResultsForModules(("WorksheetForTest", code, ComponentType.DocObject));

            Assert.IsFalse(inspectionResults.Any());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new PublicEnumerationDeclaredWithinWorksheetInspection(state, state.ProjectsProvider);
        }
    }
}
