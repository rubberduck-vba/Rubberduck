using System;
using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Resources;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class IllegalAnnotationsInspectionTests
    {
        [Test]
        [Category("Inspections")]
        public void NoAnnotation_NoResult()
        {
            const string inputCode =
                @"Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IllegalAnnotationInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void GivenLegalModuleAnnotation_NoResult()
        {
            const string inputCode = @"
Option Explicit
'@PredeclaredId
";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IllegalAnnotationInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void GivenOneIlegalModuleAnnotationAcrossModules_OneResult()
        {
            const string inputCode1 = @"
Option Explicit
'@Folder(""Legal"")

Sub DoSomething()
'@Folder(""Illegal"")
End Sub
";
            const string inputCode2 = @"
Option Explicit
'@Folder(""Legal"")
";
            var vbe = MockVbeBuilder.BuildFromStdModules(("Module1", inputCode1), ("Module2", inputCode2));
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IllegalAnnotationInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void GivenTestModule_NoResult()
        {
            const string inputCode = @"
Option Explicit

Option Private Module

'@TestModule
'@Folder(""Tests"")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject(""Rubberduck.AssertClass"")
    Set Fakes = CreateObject(""Rubberduck.FakesProvider"")
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Public Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Public Sub TestCleanup()
    'this method runs after every test in the module.
End Sub
";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IllegalAnnotationInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void SingleFolderAnnotation_NoResult()
        {
            const string inputCode =
                @"'@Folder ""Foo""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IllegalAnnotationInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void MultipleFolderAnnotations_ReturnsResult()
        {
            const string inputCode =
                @"Option Explicit
'@Folder ""Foo.Bar""
'@Folder ""Biz.Buz""
Public Sub Foo()
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IllegalAnnotationInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void CorrectTestModuleAnnotation_NoResult()
        {
            const string inputCode = @"
Option Explicit
Option Private Module

'@TestModule
'@Folder(""Tests"")

Private Assert As Object
Private Fakes As Object

Public Sub Test1()
End Sub
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IllegalAnnotationInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void IllegalTestModuleAnnotation_OneResult()
        {
            const string inputCode = @"
Option Explicit
Option Private Module

Public Sub Test1()
End Sub

'@TestModule
Public Sub Test2()
End Sub
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IllegalAnnotationInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ModuleAnnotationOnTopMostMember_NoResult()
        {
            const string inputCode = @"
Option Explicit
Option Private Module

'@TestModule
Public Sub Test1()
End Sub

Public Sub Test2()
End Sub
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IllegalAnnotationInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void InspectionType()
        {
            var inspection = new IllegalAnnotationInspection(null);
            Assert.AreEqual(CodeInspectionType.RubberduckOpportunities, inspection.InspectionType);
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "IllegalAnnotationInspection";
            var inspection = new IllegalAnnotationInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}