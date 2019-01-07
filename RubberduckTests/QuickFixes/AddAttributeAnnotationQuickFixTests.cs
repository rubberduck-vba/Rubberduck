using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class AddAttributeAnnotationQuickFixTests : QuickFixTestBase
    {
        [Test]
        [Category("QuickFixes")]
        public void KnownModuleAttributeWithoutAnnotation_QuickFixWorks()
        {
            const string inputCode =
                @"Attribute VB_PredeclaredID = True
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";
            //The attribute not present in the code pane code in the VBE.
            //So adding on top is OK.
            const string expectedCode =
                @"'@PredeclaredId
Attribute VB_PredeclaredID = True
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new MissingModuleAnnotationInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void UnknownModuleAttributeWithoutAnnotation_QuickFixWorks()
        {
            const string inputCode =
                @"Attribute VB_Ext_Key = ""Key"", ""Value""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";
            //The attribute not present in the code pane code in the VBE.
            //So adding on top is OK.
            const string expectedCode =
                @"'@ModuleAttribute VB_Ext_Key, ""Key"", ""Value""
Attribute VB_Ext_Key = ""Key"", ""Value""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new MissingModuleAnnotationInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void KnownModuleAttributeWithoutAnnotationWhileOtherAttributeWithAnnotationPresent_QuickFixWorks()
        {
            const string inputCode =
                @"Attribute VB_PredeclaredID = True
Attribute VB_Exposed = True
'@Exposed
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";
            //The attribute not present in the code pane code in the VBE.
            //So adding on top is OK.
            const string expectedCode =
                @"'@PredeclaredId
Attribute VB_PredeclaredID = True
Attribute VB_Exposed = True
'@Exposed
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new MissingModuleAnnotationInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void KnownMemberAttributeWithoutAnnotation_QuickFixWorks()
        {
            const string inputCode =
                @"
Public Sub Foo()
Attribute Foo.VB_Description = ""Desc""
    Const const1 As Integer = 9
End Sub";

            const string expectedCode =
                @"
'@Description ""Desc""
Public Sub Foo()
Attribute Foo.VB_Description = ""Desc""
    Const const1 As Integer = 9
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new MissingMemberAnnotationInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void UnknownMemberAttributeWithoutAnnotation_QuickFixWorks()
        {
            const string inputCode =
                @"
Public Sub Foo()
Attribute Foo.VB_Ext_Key = ""Key"", ""Value""
    Const const1 As Integer = 9
End Sub";

            const string expectedCode =
                @"
'@MemberAttribute VB_Ext_Key, ""Key"", ""Value""
Public Sub Foo()
Attribute Foo.VB_Ext_Key = ""Key"", ""Value""
    Const const1 As Integer = 9
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new MissingMemberAnnotationInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void KnownMemberAttributeWithoutAnnotationWhileOtherAttributeWithAnnotationPresent_QuickFixWorks()
        {
            const string inputCode =
                @"'@DefaultMember
Public Sub Foo()
Attribute Foo.VB_Description = ""Desc""
Attribute Foo.VB_UserMemId = 0
    Const const1 As Integer = 9
End Sub";

            const string expectedCode =
                @"'@DefaultMember
'@Description ""Desc""
Public Sub Foo()
Attribute Foo.VB_Description = ""Desc""
Attribute Foo.VB_UserMemId = 0
    Const const1 As Integer = 9
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new MissingMemberAnnotationInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new AddAttributeAnnotationQuickFix(new AnnotationUpdater(), new AttributeAnnotationProvider());
        }
    }
}