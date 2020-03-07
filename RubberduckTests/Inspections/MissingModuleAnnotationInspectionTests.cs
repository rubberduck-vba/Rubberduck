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
    public class MissingModuleAnnotationInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void NoAttribute_NoResult()
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
        public void NoModuleAttribute_NoResult()
        {
            const string inputCode =
                @"
Public Sub Foo()
Attribute Foo.VB_Description = ""Desc""
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void ModuleAttributeWithoutAnnotationReturnsResult()
        {
            const string inputCode =
                @"Attribute VB_Description = ""Desc""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void ModuleAttributeWithOtherValueDoesNotReturnResult()
        {
            const string inputCode =
                @"Attribute VB_Description = ""NotDesc""
'@ModuleAttribute VB_Description, ""Desc""
Public Sub Foo()
Attribute Foo.VB_UserMemId = 40
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void ModuleAttributeWithoutAnnotationInDocumentModuleDoesNotReturnResult()
        {
            const string inputCode =
                @"Attribute VB_Description = ""Desc""
Public Sub Foo()
Attribute Foo.VB_UserMemId = 40
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode, ComponentType.Document);
            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void VbExtKeyModuleAttributeWithOtherKeyReturnsResult()
        {
            const string inputCode =
                @"Attribute VB_Ext_Key = ""OtherKey"", ""SomeValue""
'@ModuleAttribute VB_Ext_Key, ""Key"", ""Value""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void VbExtKeyModuleAttributeWithOtherKeyInDocumentModuleDoesNotReturnResult()
        {
            const string inputCode =
                @"Attribute VB_Ext_Key = ""OtherKey"", ""SomeValue""
'@ModuleAttribute VB_Ext_Key, ""Key"", ""Value""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode, ComponentType.Document);
            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void VbExtKeyModuleAttributeWithSameKeyDoesNotReturnResult()
        {
            const string inputCode =
                @"Attribute VB_Ext_Key = ""Key"", ""SomeValue""
'@ModuleAttribute VB_Ext_Key, ""Key"", ""Value""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void MemberAttributeWithoutAnnotationAndIgnoreModuleDoesNotReturnResult()
        {
            const string inputCode =
                @"Attribute VB_Description = ""Desc""
'@IgnoreModule MissingModuleAnnotation
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void NameAttributeInStandardModuleDoesNotReturnResult()
        {
            const string inputCode =
                @"Attribute VB_Name = ""SomeName""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode, ComponentType.StandardModule);
            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void NameAttributeInClassModuleDoesNotReturnResult()
        {
            const string inputCode =
                @"Attribute VB_Name = ""SomeName""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode, ComponentType.ClassModule);
            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void NameAttributeInUserFormDoesNotReturnResult()
        {
            const string inputCode =
                @"Attribute VB_Name = ""SomeName""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode, ComponentType.UserForm);
            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void GlobalNameSpaceAttributeWithValueFalseInClassModuleDoesNotReturnResult()
        {
            const string inputCode =
                @"Attribute VB_GlobalNameSpace = False
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode, ComponentType.ClassModule);
            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void GlobalNameSpaceAttributeWithValueFalseInUserFormModuleDoesNotReturnResult()
        {
            const string inputCode =
                @"Attribute VB_GlobalNameSpace = False
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode, ComponentType.UserForm);
            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void GlobalNameSpaceAttributeWithValueTrueInClassModuleReturnsResult()
        {
            const string inputCode =
                @"Attribute VB_GlobalNameSpace = True
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode, ComponentType.ClassModule);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void GlobalNameSpaceAttributeWithValueTrueInUserFormReturnsResult()
        {
            const string inputCode =
                @"Attribute VB_GlobalNameSpace = True
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode, ComponentType.UserForm);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void ExposedAttributeWithValueFalseInClassModuleDoesNotReturnResult()
        {
            const string inputCode =
                @"Attribute VB_Exposed = False
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode, ComponentType.ClassModule);
            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void ExposedAttributeWithValueFalseInUserFormModuleDoesNotReturnResult()
        {
            const string inputCode =
                @"Attribute VB_Exposed = False
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode, ComponentType.UserForm);
            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void ExposedAttributeWithValueTrueInClassModuleReturnsResult()
        {
            const string inputCode =
                @"Attribute VB_Exposed = True
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode, ComponentType.ClassModule);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void ExposedAttributeWithValueTrueInUserFormReturnsResult()
        {
            const string inputCode =
                @"Attribute VB_Exposed = True
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode, ComponentType.UserForm);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void CreatableAttributeWithValueFalseInClassModuleDoesNotReturnResult()
        {
            const string inputCode =
                @"Attribute VB_Creatable = False
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode, ComponentType.ClassModule);
            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void CreatableAttributeWithValueFalseInUserFormModuleDoesNotReturnResult()
        {
            const string inputCode =
                @"Attribute VB_Creatable = False
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode, ComponentType.UserForm);
            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void CreatableAttributeWithValueTrueInClassModuleReturnsResult()
        {
            const string inputCode =
                @"Attribute VB_Creatable = True
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode, ComponentType.ClassModule);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void CreatableAttributeWithValueTrueInUserFormReturnsResult()
        {
            const string inputCode =
                @"Attribute VB_Creatable = True
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode, ComponentType.UserForm);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void PredeclaredIdAttributeWithValueFalseInClassModuleDoesNotReturnResult()
        {
            const string inputCode =
                @"Attribute VB_PredeclaredId = False
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode, ComponentType.ClassModule);
            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void PredeclaredIdAttributeWithValueFalseInUserFormModuleReturnsResult()
        {
            const string inputCode =
                @"Attribute VB_PredeclaredId = False
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode, ComponentType.UserForm);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void PredeclaredIdAttributeWithValueTrueInClassModuleReturnsResult()
        {
            const string inputCode =
                @"Attribute VB_PredeclaredId = True
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode, ComponentType.ClassModule);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void PredeclaredIdAttributeWithValueTrueInUserFormDoesNotReturnResult()
        {
            const string inputCode =
                @"Attribute VB_PredeclaredId = True
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode, ComponentType.UserForm);
            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionResultContainsAttributeNameAndValues()
        {
            const string inputCode =
                @"Attribute VB_Description = ""Desc""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            var expectedAttributeName = "VB_Description";
            var expectedAttributeValues = new List<string> { "\"Desc\"" };

            var inspectionResult = inspectionResults.Single();

            if (inspectionResult is IWithInspectionResultProperties<(string AttributeName, IReadOnlyList<string> AttributeValues)> resultProperties)
            {
                var (actualAttributeBaseName, actualAttributeValues) = resultProperties.Properties;

                Assert.AreEqual(expectedAttributeName, actualAttributeBaseName);
                Assert.AreEqual(expectedAttributeValues.Count, actualAttributeValues.Count);
                Assert.AreEqual(expectedAttributeValues[0], actualAttributeValues[0]);
            }
            else
            {
                Assert.Fail("Result is missing expected properties.");
            }
        }

        private IEnumerable<IInspectionResult> InspectionResults(string inputCode, ComponentType componentType = ComponentType.StandardModule)
            => InspectionResultsForModules(("TestModule", inputCode, componentType));

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new MissingModuleAnnotationInspection(state);
        }
    }
}