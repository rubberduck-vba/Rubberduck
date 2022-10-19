using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class PublicImplementationShouldBePrivateInspectionTests : InspectionTestsBase
    {
        [TestCase("Class_Initialize", "Private", 0)]
        [TestCase("Class_Initialize", "Public", 1)]
        [TestCase("Class_Terminate", "Friend", 1)]
        [TestCase("Class_NamedPoorly", "Friend", 0)]
        [TestCase("Class_Initialize_Again", "Friend", 0)]
        [Category("Inspections")]
        [Category(nameof(PublicImplementationShouldBePrivateInspectionTests))]
        public void LifecycleHandlers(string memberIdentifier, string accessibility, long expected)
        {
            var inputCode = 
$@"Option Explicit

Private mVal As Long

{accessibility} Sub {memberIdentifier}()
    mVal = 5
End Sub
";

            var inspectionResults = InspectionResultsForModules(
                (MockVbeBuilder.TestModuleName, inputCode, ComponentType.ClassModule)); ;

            var actual = inspectionResults.Count();

            Assert.AreEqual(expected, actual);
        }

        [TestCase("Public", 1)]
        [TestCase("Private", 0)]
        [Category("Inspections")]
        [Category(nameof(PublicImplementationShouldBePrivateInspectionTests))]
        public void UserDefinedEventHandlers(string accessibility, long expected)
        {
            var eventDeclaringClassName = "EventClass";
            var eventDeclarationCode =
$@"
    Option Explicit

    Public Event MyEvent(ByVal arg1 As Integer, ByVal arg2 As String)
";

            var inputCode =
$@"
    Option Explicit

    Private WithEvents abc As {eventDeclaringClassName}

    {accessibility} Sub abc_MyEvent(ByVal i As Integer, ByVal s As String)
    End Sub
";

            var inspectionResults = InspectionResultsForModules(
                (eventDeclaringClassName, eventDeclarationCode, ComponentType.ClassModule),
                (MockVbeBuilder.TestModuleName, inputCode, ComponentType.ClassModule));

            var actual = inspectionResults.Count();

            Assert.AreEqual(expected, actual);
        }

        [TestCase("Public", 1)]
        [TestCase("Private", 0)]
        [Category("Inspections")]
        [Category(nameof(PublicImplementationShouldBePrivateInspectionTests))]
        public void InterfaceImplementingMembers(string accessibility, long expected)
        {
            var interfaceDeclarationClass = "ITestClass";
            var interfaceDeclarationCode =
$@"
    Option Explicit

    Public Sub ImplementMe(ByVal arg1 As Integer, ByVal arg2 As String)
    End Sub
";

            var inputCode =
$@"
    Option Explicit

    Implements {interfaceDeclarationClass}

    {accessibility} Sub {interfaceDeclarationClass}_ImplementMe(ByVal i As Integer, ByVal s As String)
    End Sub
";

            var inspectionResults = InspectionResultsForModules(
                (interfaceDeclarationClass, interfaceDeclarationCode, ComponentType.ClassModule),
                (MockVbeBuilder.TestModuleName, inputCode, ComponentType.ClassModule));

            var actual = inspectionResults.Count();

            Assert.AreEqual(expected, actual);
        }

        [TestCase("Workbook_Open", "ThisWorkbook", 1)]
        [TestCase("Worksheet_SelectionChange", "Sheet1", 1)]
        [TestCase("Document_Open", "ThisDocument", 1)]
        [TestCase("Document_Open_Again", "ThisDocument", 0)]
        [Category("Inspections")]
        [Category(nameof(PublicImplementationShouldBePrivateInspectionTests))]
        public void DocumentEventHandlers(string subroutineName, string objectName, long expected)
        {
            var inputCode =
$@"
Public Sub {subroutineName}()    
    Range(""A1"").Value = ""Test""
End Sub";

            var inspectionResults = InspectionResultsForModules(
                (objectName, inputCode, ComponentType.Document)); ;

            var actual = inspectionResults.Count();

            Assert.AreEqual(expected, actual);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new PublicImplementationShouldBePrivateInspection(state);
        }
    }
}
