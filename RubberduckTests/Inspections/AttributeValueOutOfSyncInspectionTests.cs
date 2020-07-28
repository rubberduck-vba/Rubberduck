using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Annotations.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;

namespace RubberduckTests.Inspections
{
    public class AttributeValueOutOfSyncInspectionTests : InspectionTestsBase
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
        public void ModuleAttributeWithOtherValueReturnsResult()
        {
            const string inputCode =
                @"Attribute VB_Exposed = False
'@ModuleAttribute VB_Exposed, True
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void ModuleAttributeWithOtherValueInDocumentModuleDoesNotReturnResult()
        {
            const string inputCode =
                @"Attribute VB_Exposed = False
'@ModuleAttribute VB_Exposed, True
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode, ComponentType.Document);
            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void ModuleAttributeWithSameValue_NoResult()
        {
            const string inputCode =
                @"Attribute VB_Exposed = True
'@ModuleAttribute VB_Exposed, True
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void VbExtKeyModuleAttributeWithOtherKey_NoResult()
        {
            const string inputCode =
                @"Attribute VB_Ext_Key = ""OtherKey"", ""SomeValue""
'@ModuleAttribute VB_Ext_Key, ""Key"", ""Value""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void VbExtKeyModuleAttributeWithSameKeyButOtherValueReturnsResult()
        {
            const string inputCode =
                @"Attribute VB_Ext_Key = ""Key"", ""OtherValue""
'@ModuleAttribute VB_Ext_Key, ""Key"", ""Value""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void VbExtKeyModuleAttributeWithSameKeyAndValue_NoResult()
        {
            const string inputCode =
                @"Attribute VB_Ext_Key = ""Key"", ""Value""
'@ModuleAttribute VB_Ext_Key, ""Key"", ""Value""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void MemberAttributeWithOtherValueReturnsResult()
        {
            const string inputCode =
                @"'@MemberAttribute VB_UserMemId, -4
Public Sub Foo()
Attribute Foo.VB_UserMemId = 40
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void MemberAttributeWithOtherValueInDocumentModuleDoesNotReturnResult()
        {
            const string inputCode =
                @"'@MemberAttribute VB_UserMemId, -4
Public Sub Foo()
Attribute Foo.VB_UserMemId = 40
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode, ComponentType.Document);
            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void MemberAttributeWithSameValue_NoResult()
        {
            const string inputCode =
                @"'@MemberAttribute VB_UserMemId, -4
Public Sub Foo()
Attribute Foo.VB_UserMemId = -4
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void VbExtKeyMemberAttributeWithOtherKey_NoResult()
        {
            const string inputCode =
                @"'@MemberAttribute VB_Ext_Key, ""Key"", ""Value""
Public Sub Foo()
Attribute Foo.VB_Ext_Key = ""OtherKey"", ""SomeValue""
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void VbExtKeyMemberAttributeWithSameKeyButOtherValueReturnsResult()
        {
            const string inputCode =
                @"'@MemberAttribute VB_Ext_Key, ""Key"", ""Value""
Public Sub Foo()
Attribute Foo.VB_Ext_Key = ""Key"", ""OtherValue""
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void VbExtKeyMemberAttributeWithSameKeyAndValue_NoResult()
        {
            const string inputCode =
                @"'@MemberAttribute VB_Ext_Key, ""Key"", ""Value""
Public Sub Foo()
Attribute Foo.VB_Ext_Key = ""Key"", ""Value""
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void ResultContainsAnnotationAndAttributeValues()
        {
            const string inputCode =
                @"'@MemberAttribute VB_UserMemId, -4
Public Sub Foo()
Attribute Foo.VB_UserMemId = 40
    Const const1 As Integer = 9
End Sub";

            var inspectionResults = InspectionResults(inputCode);
            var inspectionResult = inspectionResults.First();

            if (inspectionResult is IWithInspectionResultProperties<(IParseTreeAnnotation Annotation, string AttributeName, IReadOnlyList<string> AttributeValues)> resultProperties)
            {
                var (pta, attributeName, attributeValues) = resultProperties.Properties;

                Assert.IsInstanceOf<MemberAttributeAnnotation>(pta.Annotation);
                Assert.AreEqual("VB_UserMemId", attributeName);
                Assert.AreEqual("-4", ((IAttributeAnnotation)pta.Annotation).AttributeValues(pta)[0]);
                Assert.AreEqual("40", attributeValues[0]);
            }
            else
            {
                Assert.Fail("Result is missing expected properties.");
            }
        }

        private IEnumerable<IInspectionResult> InspectionResults(string inputCode, ComponentType componentType = ComponentType.StandardModule)
            => InspectionResultsForModules(("TestComponent", inputCode, componentType));

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new AttributeValueOutOfSyncInspection(state);
        }
    }
}