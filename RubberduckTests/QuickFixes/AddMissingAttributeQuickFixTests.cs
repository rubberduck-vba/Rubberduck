using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class AddMissingAttributeQuickFixTests : QuickFixTestBase
    {
        [Test]
        [Category("QuickFixes")]
        public void MissingModuleAttribute_QuickFixWorks()
        {
            const string inputCode =
                @"'@ModuleAttribute VB_Description, ""Desc""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            const string expectedCode =
                @"Attribute VB_Description = ""Desc""
'@ModuleAttribute VB_Description, ""Desc""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new MissingAttributeInspection(state), CodeKind.AttributesCode);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void MissingMemberAttribute_QuickFixWorks()
        {
            const string inputCode =
                @"'@MemberAttribute VB_Description, ""Desc""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            const string expectedCode =
                @"'@MemberAttribute VB_Description, ""Desc""
Public Sub Foo()
Attribute Foo.VB_Description = ""Desc""
    Const const1 As Integer = 9
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new MissingAttributeInspection(state), CodeKind.AttributesCode);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void MissingModuleAttributeWithMultipleValues_QuickFixWorks()
        {
            const string inputCode =
                @"'@ModuleAttribute VB_Ext_Key, ""Key"", ""Value""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            const string expectedCode =
                @"Attribute VB_Ext_Key = ""Key"", ""Value""
'@ModuleAttribute VB_Ext_Key, ""Key"", ""Value""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new MissingAttributeInspection(state), CodeKind.AttributesCode);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void MissingMemberAttributeWithMultipleValues_QuickFixWorks()
        {
            const string inputCode =
                @"'@MemberAttribute VB_Ext_Key, ""Key"", ""Value""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            const string expectedCode =
                @"'@MemberAttribute VB_Ext_Key, ""Key"", ""Value""
Public Sub Foo()
Attribute Foo.VB_Ext_Key = ""Key"", ""Value""
    Const const1 As Integer = 9
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new MissingAttributeInspection(state), CodeKind.AttributesCode);
            Assert.AreEqual(expectedCode, actualCode);
        }

        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new AddMissingAttributeQuickFix(new AttributesUpdater(state));
        }
    }
}