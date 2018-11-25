using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
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
        public void FirstMemberAnnotation_IsNotIllegal_InMultipleModules()
        {
            const string inputCode1 =
                @"'@TestModule
'@Folder(""Test"")
Option Explicit

'@ModuleInitialize
Public Sub ModuleInitializeLegal()
End Sub";
            const string inputCode2 =
                @"'@TestModule
'@Folder(""Test"")
Option Explicit

'@ModuleInitialize
Public Sub ModuleInitializeAlsoLegal()
End Sub";

            var vbe = MockVbeBuilder.BuildFromStdModules(("Module1", inputCode1), ("Module2", inputCode2));
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
        public void NonExistentModuleAnnotation_OneResult()
        {
            const string inputCode = @"
'@ThisDoesNotExist
Option Explicit
Option Private Module

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

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void NonExistentMemberAnnotation_OneResult()
        {
            const string inputCode = @"
Option Explicit
Option Private Module

Public Sub Test1()
End Sub

'@ThisDoesNotExist
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
        public void ModuleAnnotationInMemberAnnotationSectionOfFirstMember_OneResult()
        {
            const string inputCode = @"
Option Explicit
Option Private Module

'@TestModule _

'@TestMethod
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

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ModuleAnnotationAboveMemberAnnotationSectionOfFirstMember_NoResult()
        {
            const string inputCode = @"
Option Explicit
Option Private Module

'@TestModule _

'
'@TestMethod
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
        public void MemberAnnotationOnTopMostMember_NoResult()
        {
            const string inputCode = @"
Option Explicit
Option Private Module

'@TestMethod _

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
        public void MemberAnnotationOnLaterMember_NoResult()
        {
            const string inputCode = @"
Option Explicit
Option Private Module

Public Sub Test1()
End Sub

'@TestMethod _

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
        public void MemberAnnotationAboveTopMostMember_OneResult()
        {
            const string inputCode = @"
Option Explicit
Option Private Module

'@TestMethod 
'
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

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void VariableAnnotationOnVariable_NoResult()
        {
            const string inputCode = @"
Option Explicit
Option Private Module

'@Obsolete
Public foo As Long

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
        public void VariableAnnotationOnModule_OneResult()
        {
            const string inputCode = @"
Option Explicit
Option Private Module

'@Obsolete 

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

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void MemberAnnotationOnVariable_OneResult()
        {
            const string inputCode = @"
Option Explicit
Option Private Module

'@TestMethod 
Public foo As Long

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
        public void MemberAnnotationOnIdentifier_OneResultPerReference()
        {
            const string inputCode = @"
Option Explicit
Option Private Module

Public foo As Long

Public Sub Test2()
    Dim a As Long
    '@TestMethod 
    a = foo
End Sub
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IllegalAnnotationInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(2, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ModuleAnnotationOnIdentifier_OneResultPerReference()
        {
            const string inputCode = @"
Option Explicit
Option Private Module

Public foo As Long

Public Sub Test2()
    Dim a As Long
    '@TestModule
    a = foo
End Sub
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IllegalAnnotationInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(2, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void IdentifierAnnotationOnIdentifier_NoResult()
        {
            const string inputCode = @"
Option Explicit
Option Private Module

Public foo As Long

Public Sub Test2()
    Dim a As Long
    '@Ignore 
    a = foo
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
        public void MemberAnnotationOnLabel_OneResult()
        {
            const string inputCode = @"
Option Explicit
Option Private Module

Public foo As Long

Public Sub Test2()
    Dim a As Long
    '@TestMethod 
label: 
    a =15
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
        public void ModuleAnnotationOnLabel_OneResult()
        {
            const string inputCode = @"
Option Explicit
Option Private Module

Public foo As Long

Public Sub Test2()
    Dim a As Long
    '@TestModule
label: 
    a =15
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
        public void VariableAnnotationOnLabel_OneResult()
        {
            const string inputCode = @"
Option Explicit
Option Private Module

Public foo As Long

Public Sub Test2()
    Dim a As Long
    '@Obsolete 
label: 
    a =15
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
        public void GeneralAnnotationOnLabel_NoResult()
        {
            const string inputCode = @"
Option Explicit
Option Private Module

Public foo As Long

Public Sub Test2()
    Dim a As Long
    '@Ignore 
label: 
    a =15
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
        public void InspectionName()
        {
            const string inspectionName = "IllegalAnnotationInspection";
            var inspection = new IllegalAnnotationInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }

        [Test]
        [Category("Inspections")]
        public void AnnotationIsCaseInsensitive()
        {
            const string inputCode =
                @"'@folder ""Foo""

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
    }
}