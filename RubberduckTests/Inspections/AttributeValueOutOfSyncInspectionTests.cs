using System.Collections.Generic;
using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    public class AttributeValueOutOfSyncInspectionTests
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
            Assert.AreEqual(AnnotationType.MemberAttribute, inspectionResult.Properties.Annotation.AnnotationType);
            Assert.AreEqual("VB_UserMemId", inspectionResult.Properties.Annotation.Attribute);
            Assert.AreEqual("-4", inspectionResult.Properties.Annotation.AttributeValues[0]);
            Assert.AreEqual("40", inspectionResult.Properties.AttributeValues[0]);
        }

        private IEnumerable<IInspectionResult> InspectionResults(string inputCode, ComponentType componentType = ComponentType.StandardModule)
        {
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, componentType, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new AttributeValueOutOfSyncInspection(state);
                return inspection.GetInspectionResults(CancellationToken.None);
            }
        }
    }
}