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
        public void VariableNotAssigned_ReturnsResult()
        {
            const string inputCode =
                @"Sub Foo()
    Dim var1 As String
End Sub";

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
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
        public void UnassignedVariable_DoesNotReturnResult()
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
