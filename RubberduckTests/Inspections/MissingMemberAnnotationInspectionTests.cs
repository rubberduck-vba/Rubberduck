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
    public class MissingMemberAnnotationInspectionTests : InspectionTestsBase
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
        public void NoMemberAttribute_NoResult()
        {
            const string inputCode =
                @"Attribute VB_Description = ""Desc""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void MemberAttributeWithoutAnnotationReturnsResult()
        {
            const string inputCode =
                @"
Public Sub Foo()
Attribute Foo.VB_UserMemId = -4
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void MemberAttributeWithOtherValueDoesNotReturnResult()
        {
            const string inputCode =
                @"'@MemberAttribute VB_UserMemId, -4
Public Sub Foo()
Attribute Foo.VB_UserMemId = 40
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void MemberAttributeWithoutAnnotationInDocumentModuleDoesNotReturnResult()
        {
            const string inputCode =
                @"
Public Sub Foo()
Attribute Foo.VB_UserMemId = 40
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode, ComponentType.Document);
            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void VbExtKeyMemberAttributeWithOtherKeyReturnsResult()
        {
            const string inputCode =
                @"'@MemberAttribute VB_Ext_Key, ""Key"", ""Value""
Public Sub Foo()
Attribute Foo.VB_Ext_Key = ""OtherKey"", ""SomeValue""
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void VbExtKeyMemberAttributeWithOtherKeyInDocumentModuleDoesNotReturnResult()
        {
            const string inputCode =
                @"'@MemberAttribute VB_Ext_Key, ""Key"", ""Value""
Public Sub Foo()
Attribute Foo.VB_Ext_Key = ""OtherKey"", ""SomeValue""
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode, ComponentType.Document);
            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void VbExtKeyMemberAttributeWithSameKeyDoesNotReturnResult()
        {
            const string inputCode =
                @"'@MemberAttribute VB_Ext_Key, ""Key"", ""Value""
Public Sub Foo()
Attribute Foo.VB_Ext_Key = ""Key"", ""SomeValue""
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void MemberAttributeWithoutAnnotationAndIgnoreDoesNotReturnResult()
        {
            const string inputCode =
                @"'@Ignore MissingMemberAnnotation
Public Sub Foo()
Attribute Foo.VB_UserMemId = -4
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
                @"'@IgnoreModule MissingMemberAnnotation
Public Sub Bar()
End Sub

Public Sub Foo()
Attribute Foo.VB_UserMemId = -4
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionResultContainsAttributeBaseNameAndValues()
        {
            const string inputCode =
                @"
Public Sub Foo()
Attribute Foo.VB_UserMemId = -4
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            var expectedAttributeName = "VB_UserMemId";
            var expectedAttributeValues = new List<string>{"-4"};

            var inspectionResult = inspectionResults.Single();
            if (inspectionResult is IWithInspectionResultProperties<(string AttributeName, IReadOnlyList<string>AttributeValues)> resultProperties)
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
            return new MissingMemberAnnotationInspection(state);
        }
    }
}