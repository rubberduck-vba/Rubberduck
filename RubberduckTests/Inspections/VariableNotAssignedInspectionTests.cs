using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class VariableNotAssignedInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void VariableNotAssigned_ReturnsResult_Local()
        {
            const string inputCode =
                @"Sub Foo()
    Dim var1 As String
End Sub";

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        [TestCase("Public")]
        [TestCase("Private")]
        public void VariableNotAssigned_ReturnsResult_Module(string scopeIdentifier)
        {
            var inputCode =
$@"
    {scopeIdentifier} Bar As Variant
";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableNotAssigned_ReturnsResult_Module_Exposed_Private()
        {
            var inputCode =
$@"
Attribute VB_Exposed = True

    Private Bar As Variant
";
            Assert.AreEqual(1, InspectionResultsForModules(("Class1", inputCode, ComponentType.ClassModule)).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableNotAssigned_DoesNotReturnResult_Module_Exposed_Public()
        {
            var inputCode =
$@"
Attribute VB_Exposed = True

    Public Bar As Variant
";
            Assert.AreEqual(0, InspectionResultsForModules(("Class1", inputCode, ComponentType.ClassModule)).Count());
        }

        [Test]
        [Category("Inspections")]
        public void UnassignedVariable_ReturnsResult_MultipleVariables()
        {
            const string inputCode =
                @"Sub Foo()
    Dim var1 As String
    Dim var2 As Date
End Sub";

            Assert.AreEqual(2, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void AssignedVariable_DoesNotReturnResult()
        {
            const string inputCode =
                @"Function Foo() As Boolean
    Dim var1 as String
    var1 = ""test""
End Function";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        [TestCase("Public")]
        [TestCase("Private")]
        public void AssignedVariable_DoesNotReturnResult_Module(string scopeIdentifier)
        {
            var inputCode =
$@"
    {scopeIdentifier} Bar As Variant

Sub Foo()
    Bar = ""test""
End Sub
";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void UnassignedVariable_ReturnsResult_MultipleVariables_SomeAssigned()
        {
            const string inputCode =
                @"Sub Foo()
    Dim var1 as Integer
    var1 = 8

    Dim var2 as String
End Sub";

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableNotAssigned_GivenByRefAssignment_DoesNotReturnResult()
        {
            const string inputCode = @"
Sub Foo()
    Dim var1 As String
    Bar var1
End Sub

Sub Bar(ByRef value As String)
    value = ""test""
End Sub
";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableNotAssigned_StrictlyInsideByRefAssignment_ReturnsResult()
        {
            const string inputCode = @"
Sub Foo()
    Dim var1 As String
    Bar var1 & ""WTF""
End Sub

Sub Bar(ByRef value As String)
    value = ""test""
End Sub
";

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableNotAssigned_ModuleVariable_ReturnsResult_Private()
        {
            const string inputCode = @"
Private myVariable As Variant

Sub Foo()
End Sub
";

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableNotAssigned_ModuleVariable_ReturnsResult_Public()
        {
            const string inputCode = @"
Public myVariable As Variant

Sub Foo()
End Sub
";

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableNotAssigned_GivenByRefAssignment_ModuleVariable_DoesNotReturnResult()
        {
            const string inputCode = @"
Public myVariable As Variant

Sub Foo()
    Bar myVariable
End Sub

Sub Bar(ByRef value As String)
    value = ""test""
End Sub
";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableNotAssigned_StrictlyInsideByRefAssignment_ModuleVariable_ReturnsResult()
        {
            const string inputCode = @"
Public myVariable As Variant

Sub Foo()
    Bar myVariable & ""WTF""
End Sub

Sub Bar(ByRef value As String)
    value = ""test""
End Sub
";

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableNotAssigned_AssignmentFromOtherModule_ModuleVariable_DoesNotReturnResult()
        {
            const string otherModuleCode = @"
Public myVariable As Variant
";

            const string moduleCode = @"
Sub Foo()
    myVariable = 42
End Sub
";

            Assert.AreEqual(0, InspectionResultsForModules(
                ("OtherModule", otherModuleCode, ComponentType.StandardModule),
                ("TestModule", moduleCode, ComponentType.StandardModule)
            ).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableNotAssigned_GivenByRefAssignmentInOtherModule_ModuleVariable_DoesNotReturnResult()
        {
            const string otherModuleCode = @"
Public myVariable As Variant
";

            const string moduleCode = @"
Sub Foo()
    Bar myVariable
End Sub

Sub Bar(ByRef value As String)
    value = ""test""
End Sub
";

            Assert.AreEqual(0, InspectionResultsForModules(
                ("OtherModule", otherModuleCode, ComponentType.StandardModule),
                ("TestModule", moduleCode, ComponentType.StandardModule)
            ).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableNotAssigned_StrictlyInsideByRefAssignmentInOtherModule_QualifiedModuleVariable_ReturnsResult()
        {
            const string otherModuleCode = @"
Public myVariable As Variant
";

            const string moduleCode = @"
Sub Foo()
    Bar myVariable & ""WTF""
End Sub

Sub Bar(ByRef value As String)
    value = ""test""
End Sub
";

            Assert.AreEqual(1, InspectionResultsForModules(
                ("OtherModule", otherModuleCode, ComponentType.StandardModule),
                ("TestModule", moduleCode, ComponentType.StandardModule)
            ).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableNotAssigned_Assignment_QualifiedModuleVariable_DoesNotReturnResult()
        {
            const string otherModuleCode = @"
Public myVariable As Variant
";

            const string moduleCode = @"
Sub Foo()
    OtherModule.myVariable = 42
End Sub
";

            Assert.AreEqual(0, InspectionResultsForModules(
                ("OtherModule", otherModuleCode, ComponentType.StandardModule),
                ("TestModule", moduleCode, ComponentType.StandardModule)
            ).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableNotAssigned_GivenByRefAssignment_QualifiedModuleVariable_DoesNotReturnResult()
        {
            const string otherModuleCode = @"
Public myVariable As Variant
";

            const string moduleCode = @"
Sub Foo()
    Bar OtherModule.myVariable
End Sub

Sub Bar(ByRef value As String)
    value = ""test""
End Sub
";

            Assert.AreEqual(0, InspectionResultsForModules(
                ("OtherModule", otherModuleCode, ComponentType.StandardModule),
                ("TestModule", moduleCode, ComponentType.StandardModule)
                ).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableNotAssigned_StrictlyInsideByRefAssignment_QualifiedModuleVariable_ReturnsResult()
        {
            const string otherModuleCode = @"
Public myVariable As Variant
";

            const string moduleCode = @"
Sub Foo()
    Bar OtherModule.myVariable & ""WTF""
End Sub

Sub Bar(ByRef value As String)
    value = ""test""
End Sub
";

            Assert.AreEqual(1, InspectionResultsForModules(
                ("OtherModule", otherModuleCode, ComponentType.StandardModule),
                ("TestModule", moduleCode, ComponentType.StandardModule)
            ).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableNotAssigned_AssignmentInOtherModule_ClassModuleVariable_DoesNotReturnResult()
        {
            const string classModuleCode = @"
Public myVariable As Variant
";

            const string moduleCode = @"
Sub Foo()
    Dim cls As TestClass
    Set cls = New TestClass
    cls.myVariable = 42
End Sub

Sub Bar(ByRef value As String)
    value = ""test""
End Sub
";

            Assert.AreEqual(0, InspectionResultsForModules(
                ("TestClass", classModuleCode, ComponentType.ClassModule),
                ("TestModule", moduleCode, ComponentType.StandardModule)
            ).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableNotAssigned_GivenByRefAssignment_ClassModuleVariable_DoesNotReturnResult()
        {
            const string classModuleCode = @"
Public myVariable As Variant
";

            const string moduleCode = @"
Sub Foo()
    Dim cls As TestClass
    Set cls = New TestClass
    Bar cls.myVariable
End Sub

Sub Bar(ByRef value As String)
    value = ""test""
End Sub
";

            Assert.AreEqual(0, InspectionResultsForModules(
                ("TestClass", classModuleCode, ComponentType.ClassModule),
                ("TestModule", moduleCode, ComponentType.StandardModule)
            ).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableNotAssigned_StrictlyInsideByRefAssignment_ClassModuleVariable_ReturnsResult()
        {
            const string classModuleCode = @"
Public myVariable As Variant
";

            const string moduleCode = @"
Sub Foo()
    Dim cls As TestClass
    Set cls = New TestClass
    Bar cls.myVariable & ""WTF""
End Sub

Sub Bar(ByRef value As String)
    value = ""test""
End Sub
";

            Assert.AreEqual(1, InspectionResultsForModules(
                ("TestClass", classModuleCode, ComponentType.ClassModule),
                ("TestModule", moduleCode, ComponentType.StandardModule)
            ).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableNotAssigned_ArrayWithElementAssignment_DoesNotReturnResult()
        {
            const string inputCode = @"
Public Sub Foo()
    Dim arr(0 To 0) As Variant
    arr(0) = 42
End Sub
";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableNotAssigned_ReDimDeclaredArrayWithElementAssignment_DoesNotReturnResult()
        {
            const string inputCode = @"
Public Sub Foo()
    ReDim arr(0 To 0) As Variant
    arr(0) = 42
End Sub
";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        //See issue #5845 at https://github.com/rubberduck-vba/Rubberduck/issues/5845
        public void VariableNotAssigned_VariantUsedAsArrayWithElementAssignment_DoesNotReturnResult()
        {
            const string inputCode = @"
Public Sub Foo()
    Dim arr As Variant
    ReDim arr(0 To 0) As Variant
    arr(0) = 42
End Sub
";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableNotAssigned_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"Sub Foo()
'@Ignore VariableNotAssigned
Dim var1 As String
End Sub";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new VariableNotAssignedInspection(null);

            Assert.AreEqual(nameof(VariableNotAssignedInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new VariableNotAssignedInspection(state);
        }
    }
}
