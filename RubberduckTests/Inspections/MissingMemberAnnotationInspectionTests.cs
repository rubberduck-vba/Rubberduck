using System.Collections.Generic;
using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class MissingMemberAnnotationInspectionTests
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
            var actualAttributeBaseName = inspectionResult.Properties.AttributeName;
            var actualAttributeValues = inspectionResult.Properties.AttributeValues;

            Assert.AreEqual(expectedAttributeName, actualAttributeBaseName);
            Assert.AreEqual(expectedAttributeValues.Count, actualAttributeValues.Count);
            Assert.AreEqual(expectedAttributeValues[0], actualAttributeValues[0]);
        }

        private IEnumerable<IInspectionResult> InspectionResults(string inputCode, ComponentType componentType = ComponentType.StandardModule)
        {
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, componentType, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new MissingMemberAnnotationInspection(state);
                return inspection.GetInspectionResults(CancellationToken.None);
            }
        }
    }
}