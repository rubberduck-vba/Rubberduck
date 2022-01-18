using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class MissingAttributeInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void NoAnnotation_NoResult()
        {
            const string inputCode =
                @"Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void ModuleAttributeAnnotationWithoutAttributeReturnsResult()
        {
            const string inputCode =
                @"'@ModuleAttribute VB_Description, ""Desc""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void ModuleAttributeAnnotationWithoutAttributeInDocumentModuleDoesNotReturnResult()
        {
            const string inputCode =
                @"'@ModuleAttribute VB_Description, ""Desc""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode, ComponentType.Document);
            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void ModuleAttributeAnnotationWithAttributeReturnsNoResult()
        {
            const string inputCode =
                @"Attribute VB_Description = ""NotDesc""
'@ModuleAttribute VB_Description, ""Desc""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void VbExtKeyModuleAttributeAnnotationWithAttributeButWithoutKeyReturnsResult()
        {
            const string inputCode =
                @"Attribute VB_Ext_Key = ""OtherKey"", ""Value""
'@ModuleAttribute VB_Ext_Key, ""Key"", ""Value""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void VbExtKeyModuleAttributeAnnotationWithAttributeAndKeyReturnsNoResult()
        {
            const string inputCode =
                @"Attribute VB_Ext_Key = ""Key"", ""OtherValue""
'@ModuleAttribute VB_Ext_Key, ""Key"", ""Value""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void ModuleAttributeAnnotationWithMissingArgumentsNoResult()
        {
            const string inputCode =
                @"'@ModuleAttribute
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void MemberAttributeAnnotationWithoutAttributeReturnsResult()
        {
            const string inputCode =
                @"'@MemberAttribute VB_Description, ""Desc""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void MemberAttributeAnnotationWithoutAttributeInDomcumentModuleDoesNotReturnResult()
        {
            const string inputCode =
                @"'@MemberAttribute VB_Description, ""Desc""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode, ComponentType.Document);
            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void MemberAttributeAnnotationWithAttributeReturnsNoResult()
        {
            const string inputCode =
                @"'@MemberAttribute VB_Description, ""Desc""
Public Sub Foo()
Attribute Foo.VB_Description = ""NotDesc""
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void VbExtKeyMemberAttributeAnnotationWithAttributeButWithoutKeyReturnsResult()
        {
            const string inputCode =
                @"'@MemberAttribute VB_Ext_Key, ""Key"", ""Value""
Public Sub Foo()
Attribute Foo.VB_Ext_Key = ""OtherKey"", ""Value""
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void VbExtKeyMemberAttributeAnnotationWithAttributeAndKeyReturnsNoResult()
        {
            const string inputCode =
                @"'@MemberAttribute VB_Ext_Key, ""Key"", ""Value""
Public Sub Foo()
Attribute Foo.VB_Ext_Key = ""Key"", ""OtherValue""
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void MemberAttributeAnnotationWithMissingArgumentsNoResult()
        {
            const string inputCode =
                @"'@MemberAttribute
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void DefaultMemberAnnotationWithoutAttributeReturnsResult()
        {
            const string inputCode =
                @"'@DefaultMember
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void DefaultMemberAnnotationWithAttributeReturnsNoResult()
        {
            const string inputCode =
                @"'@DefaultMember
Public Sub Foo()
Attribute Foo.VB_UserMemId = 42
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void EnumeratorAnnotationWithoutAttributeReturnsResult()
        {
            const string inputCode =
                @"'@Enumerator
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void EnumeratorAnnotationWithAttributeReturnsNoResult()
        {
            const string inputCode =
                @"'@Enumerator
Public Sub Foo()
Attribute Foo.VB_UserMemId = 42
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void DescriptionAnnotationWithoutAttributeReturnsResult()
        {
            const string inputCode =
                @"'@Description ""Desc""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void DescriptionAnnotationWithAttributeReturnsNoResult()
        {
            const string inputCode =
                @"'@Description, ""Desc""
Public Sub Foo()
Attribute Foo.VB_Description = ""NotDesc""
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void VariableDescriptionAnnotationWithoutAttributeReturnsResult_Variable()
        {
            const string inputCode =
                @"'@VariableDescription ""Desc""
Public Foo As String
";

            var inspectionResults = InspectionResults(inputCode);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void Variable_DescriptionAnnotationWithAttributeReturnsNoResult_Variable()
        {
            const string inputCode =
                @"'@VariableDescription ""Desc""
Public Foo As String
Attribute Foo.VB_VarDescription = ""NotDesc""
";

            var inspectionResults = InspectionResults(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void VariableDescriptionAnnotationWithoutAttributeReturnsResult_Constant()
        {
            const string inputCode =
                @"'@VariableDescription ""Desc""
Public Const Foo As String = ""Huh""
";

            var inspectionResults = InspectionResults(inputCode);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void Variable_DescriptionAnnotationWithAttributeReturnsNoResult_Constant()
        {
            const string inputCode =
                @"'@VariableDescription ""Desc""
Public Const Foo As String = ""Huh""
Attribute Foo.VB_VarDescription = ""NotDesc""
";

            var inspectionResults = InspectionResults(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void ModuleDescriptionAnnotationWithoutAttributeReturnsResult()
        {
            const string inputCode =
                @"'@ModuleDescription ""Desc""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void ModuleDescriptionAnnotationWithAttributeReturnsNoResult()
        {
            const string inputCode =
                @"Attribute VB_Description = ""NotDesc""
'@ModuleDescription ""Desc""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void PredeclaredIdAnnotationWithoutAttributeReturnsResult()
        {
            const string inputCode =
                @"'@PredeclaredId
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void PredeclaredIdAnnotationWithAttributeReturnsNoResult()
        {
            const string inputCode =
                @"Attribute VB_PredeclaredId = False
'@PredeclaredId
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void ExposedAnnotationWithoutAttributeReturnsResult()
        {
            const string inputCode =
                @"'@Exposed
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void ExposedAnnotationWithAttributeReturnsNoResult()
        {
            const string inputCode =
                @"Attribute VB_Exposed = False
'@Exposed
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void AttributeAnnotationOnDeclarationNotAllowingAttributes_NoResult()
        {
            const string inputCode =
                @"
Private Sub Foo()
'local variables do not allow attributes
    '@VariableDescription(""Desc"")
    Dim bar As Variant
End Sub
";

            var inspectionResults = InspectionResults(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void MissingMemberAttributeOnDeclareStatement_OneResult()
        {
            const string inputCode =
                @"'@Description(""Desc"")
Private Declare Sub CopyMemory Lib ""kernel32.dll"" Alias ""RtlMoveMemory"" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)
";

            var inspectionResults = InspectionResults(inputCode);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void MemberAttributeAnnotationOnDeclareStatement_WithAttribute_NoResult()
        {
            const string inputCode =
                @"'@Description(""Desc"")
Private Declare Sub CopyMemory Lib ""kernel32.dll"" Alias ""RtlMoveMemory"" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)
Attribute CopyMemory.VB_Description = ""Desc""
";

            var inspectionResults = InspectionResults(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }


        private IEnumerable<IInspectionResult> InspectionResults(string inputCode, ComponentType componentType = ComponentType.StandardModule)
        {
            return InspectionResultsForModules(("TestModule", inputCode, componentType));
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new MissingAttributeInspection(state);
        }
    }
}