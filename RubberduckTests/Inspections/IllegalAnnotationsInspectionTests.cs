using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class AnnotationInIncompatibleComponentTypeInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void ModuleAttributeAnnotationInDocumentReturnsResult()
        {
            const string inputCode = @"
'@ModuleAttribute VB_Description, ""Desc""

";

            var inspectionResults = InspectionResultsForModules(("TestDocument", inputCode, ComponentType.Document));
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void MemberAttributeAnnotationInDocumentReturnsResult()
        {
            const string inputCode = @"
'@MemberAttribute VB_Description, ""Desc""
Public Sub Foo()
End Sub
";

            var inspectionResults = InspectionResultsForModules(("TestDocument", inputCode, ComponentType.Document));
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = nameof(AnnotationInIncompatibleComponentTypeInspection);
            var inspection = InspectionUnderTest(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state) => new AnnotationInIncompatibleComponentTypeInspection(state);
    }

    [TestFixture]
    public class UnrecognizedAnnotationInspectionTests : InspectionTestsBase
    {
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

            var inspectionResults = InspectionResultsForStandardModule(inputCode);
            Assert.AreEqual(1, inspectionResults.Count());
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

            var inspectionResults = InspectionResultsForStandardModule(inputCode);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = nameof(UnrecognizedAnnotationInspection);
            var inspection = InspectionUnderTest(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state) => new UnrecognizedAnnotationInspection(state);
    }

    [TestFixture]
    public class InvalidAnnotationsInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void NoAnnotation_NoResult()
        {
            const string inputCode =
                @"Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResultsForStandardModule(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void AttributeAnnotationOnDeclarationNotAllowingAttributes_OneResult()
        {
            const string inputCode =
                @"
Private Sub Foo()
'local variables do not allow attributes
    '@VariableDescription(""Desc"")
    Dim bar As Variant
End Sub
";

            var inspectionResults = InspectionResultsForStandardModule(inputCode);
            Assert.AreEqual(1, inspectionResults.Count());
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

            var inspectionResults = InspectionResultsForModules(
                ("Module1", inputCode1, ComponentType.StandardModule), 
                ("Module2", inputCode2, ComponentType.StandardModule));

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void GivenLegalModuleAnnotation_NoResult()
        {
            const string inputCode = @"
Option Explicit
'@PredeclaredId
";

            var inspectionResults = InspectionResultsForStandardModule(inputCode);
            Assert.IsFalse(inspectionResults.Any());
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

            var inspectionResults = InspectionResultsForModules(
                ("Module1", inputCode1, ComponentType.StandardModule),
                ("Module2", inputCode2, ComponentType.StandardModule));

            Assert.AreEqual(1, inspectionResults.Count());
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

            var inspectionResults = InspectionResultsForStandardModule(inputCode);
            Assert.IsFalse(inspectionResults.Any());
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

            var inspectionResults = InspectionResultsForStandardModule(inputCode);
            Assert.IsFalse(inspectionResults.Any());
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

            var inspectionResults = InspectionResultsForStandardModule(inputCode);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void ModuleAnnotationEndingMemberAnnotationSectionOfFirstMember_NoResult()
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

            var inspectionResults = InspectionResultsForStandardModule(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void MemberAnnotationAboveModuleAnnotationEndingMemberAnnotationSectionOfFirstMember_NoResult()
        {
            const string inputCode = @"
Option Explicit
Option Private Module

'@Description ""Test""
'@TestModule _

'@TestMethod
Public Sub Test1()
End Sub

Public Sub Test2()
End Sub
";

            var inspectionResults = InspectionResultsForStandardModule(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void MemberAnnotationOnMember_OneResult()
        {
            const string inputCode = @"
Option Explicit
Option Private Module

'@TestModule _

'@TestMethod
Public Sub Test1() '@Description ""Test""
End Sub

Public Sub Test2()
End Sub
";

            var inspectionResults = InspectionResultsForStandardModule(inputCode);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void MemberAnnotationBelowLastMember_OneResult()
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
'@Description ""Test""
";

            var inspectionResults = InspectionResultsForStandardModule(inputCode);
            Assert.AreEqual(1, inspectionResults.Count());
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

            var inspectionResults = InspectionResultsForStandardModule(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void MemberAnnotationRightAboveTopMostMember_NoResult()
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

            var inspectionResults = InspectionResultsForStandardModule(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void MemberAnnotationAboveLaterMember_NoResult()
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

            var inspectionResults = InspectionResultsForStandardModule(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void MemberAnnotationAboveTopMostMember_NoResult()
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

            var inspectionResults = InspectionResultsForStandardModule(inputCode);
            Assert.IsFalse(inspectionResults.Any());
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

            var inspectionResults = InspectionResultsForStandardModule(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void VariableAnnotationOnConstant_NoResult()
        {
            const string inputCode = @"
Option Explicit
Option Private Module

'@VariableDescription ""Test""
Private Const Test As String = ""TestTestTest""

Public Sub Test2()
End Sub
";

            var inspectionResults = InspectionResultsForStandardModule(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void VariableOnOrAboveNonWhitespaceAboveFirstVariable_OneResultEach()
        {
            const string inputCode = @"'@Obsolete 
Option Explicit
Option Private Module _
'@Obsolete 

Public Sub Test1()
End Sub

Public Sub Test2()
End Sub
";

            var inspectionResults = InspectionResultsForStandardModule(inputCode);
            Assert.AreEqual(2, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void ModuleAnnotationAboveVariableAnnotationSection_NoResult()
        {
            const string inputCode = @"
Option Explicit
Option Private Module

'@TestModule
'@Obsolete
Public foo As Long

Public Sub Test2()
End Sub
";

            var inspectionResults = InspectionResultsForStandardModule(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void VariableAnnotationAboveModuleAnnotationAboveVariableAnnotation_NoResult()
        {
            const string inputCode = @"
Option Explicit
Option Private Module

'@Obsolete 
'@TestModule
Public Sub Test1()
End Sub

Public Sub Test2()
End Sub
";

            var inspectionResults = InspectionResultsForStandardModule(inputCode);
            Assert.IsFalse(inspectionResults.Any());
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

            var inspectionResults = InspectionResultsForStandardModule(inputCode);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void MemberAnnotationOnIdentifier_OneResult()
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

            var inspectionResults = InspectionResultsForStandardModule(inputCode);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void ModuleAnnotationOnIdentifierBelowDeclarationsSection_OneResult()
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

            var inspectionResults = InspectionResultsForStandardModule(inputCode);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void IdentifierAnnotationAboveIdentifier_NoResult()
        {
            const string inputCode = @"
Option Explicit
Option Private Module

Public foo As Long

Public Sub Test2()
    Dim a As Long

    '@Ignore 
    'Some comment

    a = foo
End Sub
";

            var inspectionResults = InspectionResultsForStandardModule(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void IdentifierAnnotationOnNonWhitespaceAboveIdentifier_OneResult()
        {
            const string inputCode = @"
Option Explicit
Option Private Module

Public foo As Long

Public Sub Test2()
    Dim a As Long '@Ignore 
    a = foo
End Sub
";

            var inspectionResults = InspectionResultsForStandardModule(inputCode);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void IdentifierAnnotationOnIdentifier_OneResult()
        {
            const string inputCode = @"
Option Explicit
Option Private Module

Public foo As Long

Public Sub Test2()
    Dim a As Long 
    a = foo '@Ignore
End Sub
";

            var inspectionResults = InspectionResultsForStandardModule(inputCode);
            Assert.AreEqual(1, inspectionResults.Count());
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

            var inspectionResults = InspectionResultsForStandardModule(inputCode);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void ModuleAnnotationOnLabelBelowDeclarationsSection_OneResult()
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

            var inspectionResults = InspectionResultsForStandardModule(inputCode);
            Assert.AreEqual(1, inspectionResults.Count());
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

            var inspectionResults = InspectionResultsForStandardModule(inputCode);
            Assert.AreEqual(1, inspectionResults.Count());
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

            var inspectionResults = InspectionResultsForStandardModule(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void GeneralAnnotationOnNonDeclarationNonIdentifier_NoResult()
        {
            const string inputCode = @"
Option Explicit
'@Ignore OptionBase
Option Base 1

Public foo As Long

Public Sub Test2()
End Sub
";

            var inspectionResults = InspectionResultsForStandardModule(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        //Issue #4558.
        [Test]
        [Category("Inspections")]
        public void FolderBelowOptionExplicitAndAboveImplements_NoResult()
        {
            const string inputCode = @"Option Explicit
'@Folder(""Excel.Abstract"")
Implements IWorkbookData

";
            const string interfaceCode = @"Option Explicit
'@Folder(""Excel.Abstract"")
Implements IWorkbookData

";

            var inspectionResults = InspectionResultsForModules(
                ("testClass", inputCode, ComponentType.ClassModule), 
                ("IWorkbookData", interfaceCode, ComponentType.ClassModule));

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = nameof(InvalidAnnotationInspection);
            var inspection = InspectionUnderTest(null);

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

            var inspectionResults = InspectionResultsForStandardModule(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state) => new InvalidAnnotationInspection(state);        
    }
}