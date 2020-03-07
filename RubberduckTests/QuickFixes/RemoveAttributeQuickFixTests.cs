using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class RemoveAttributeQuickFixTests : QuickFixTestBase
    {
        [Test]
        [Category("QuickFixes")]
        public void ModuleAttributeWithoutAnnotation_QuickFixWorks()
        {
            const string inputCode =
                @"Attribute VB_Description = ""Desc""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            const string expectedCode =
                @"Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new MissingModuleAnnotationInspection(state), CodeKind.AttributesCode);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void VbExtKeyModuleAttributeWithoutAnnotationForOneKey_QuickFixWorks()
        {
            const string inputCode =
                @"Attribute VB_Ext_Key = ""Key"", ""NotValue""
Attribute VB_Ext_Key = ""OtherKey"", ""OtherValue""
'@ModuleAttribute VB_Ext_Key, ""Key"", ""Value""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            const string expectedCode =
                @"Attribute VB_Ext_Key = ""Key"", ""NotValue""
'@ModuleAttribute VB_Ext_Key, ""Key"", ""Value""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new MissingModuleAnnotationInspection(state), CodeKind.AttributesCode);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void MemberAttributeWithoutAnnotation_QuickFixWorks()
        {
            const string inputCode =
                @"
Public Sub Foo()
Attribute Foo.VB_Description = ""NotDesc""
    Const const1 As Integer = 9
End Sub";

            const string expectedCode =
                @"
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new MissingMemberAnnotationInspection(state), CodeKind.AttributesCode);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void VbExtKeyMemberAttributeWithoutAnnotationForOneKey_QuickFixWorks()
        {
            const string inputCode =
                @"'@MemberAttribute VB_Ext_Key, ""Key"", ""Value""
Public Sub Foo()
Attribute Foo.VB_Ext_Key = ""Key"", ""NotValue""
Attribute Foo.VB_Ext_Key = ""OtherKey"", ""OtherValue""
    Const const1 As Integer = 9
End Sub";

            const string expectedCode =
                @"'@MemberAttribute VB_Ext_Key, ""Key"", ""Value""
Public Sub Foo()
Attribute Foo.VB_Ext_Key = ""Key"", ""NotValue""
    Const const1 As Integer = 9
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new MissingMemberAnnotationInspection(state), CodeKind.AttributesCode);
            Assert.AreEqual(expectedCode, actualCode);
        }

        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new RemoveAttributeQuickFix(new AttributesUpdater(state));
        }
    }
}