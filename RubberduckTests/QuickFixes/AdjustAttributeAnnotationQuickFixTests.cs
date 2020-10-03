using NUnit.Framework;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using System;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class AdjustAttributeAnnotationQuickFixTests : QuickFixTestBase
    {
        [Test]
        [Category("QuickFixes")]
        public void ModuleAttributeWithDifferentNonStandardValue_QuickFixWorks()
        {
            const string inputCode =
                @"Attribute VB_Description = ""OtherDesc""
'@ModuleDescription ""Desc""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            const string expectedCode =
                @"Attribute VB_Description = ""OtherDesc""
'@ModuleDescription ""OtherDesc""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new AttributeValueOutOfSyncInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ModuleAttributeWithDifferingStandardValue_RemovesAnnotation()
        {
            const string inputCode =
                @"Attribute VB_PredeclaredId = False
'@PredeclaredId
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            const string expectedCode =
                @"Attribute VB_PredeclaredId = False
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new AttributeValueOutOfSyncInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void MemberAttributeWithDifferingKnownValue_QuickFixWorks()
        {
            const string inputCode =
                @"'@DefaultMember
Public Sub Foo()
Attribute Foo.VB_UserMemId = -4
    Const const1 As Integer = 9
End Sub";

            const string expectedCode =
                @"'@Enumerator
Public Sub Foo()
Attribute Foo.VB_UserMemId = -4
    Const const1 As Integer = 9
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new AttributeValueOutOfSyncInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void MemberAttributeWithDifferingUnknownValue_QuickFixWorks()
        {
            const string inputCode =
                @"'@DefaultMember
Public Sub Foo()
Attribute Foo.VB_UserMemId = -40
    Const const1 As Integer = 9
End Sub";

            const string expectedCode =
                @"'@MemberAttribute VB_UserMemId, -40
Public Sub Foo()
Attribute Foo.VB_UserMemId = -40
    Const const1 As Integer = 9
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new AttributeValueOutOfSyncInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        protected override IVBE TestVbe(string code, out IVBComponent component)
        {
            return MockVbeBuilder.BuildFromSingleModule(code, ComponentType.ClassModule, out component).Object;
        }

        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new AdjustAttributeAnnotationQuickFix(new AnnotationUpdater(state), 
                new AttributeAnnotationProvider(MockParser.WellKnownAnnotations().OfType<IAttributeAnnotation>()));
        }
    }
}