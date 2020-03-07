using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;

namespace RubberduckTests.QuickFixes
{
    public class AdjustAttributeValuesQuickFixTests : QuickFixTestBase
    {
        [Test]
        [Category("QuickFixes")]
        public void ModuleAttributeOutOfSync_QuickFixWorks()
        {
            const string inputCode =
                @"Attribute VB_Description = ""NotDesc""
'@ModuleAttribute VB_Description, ""Desc""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            const string expectedCode =
                @"Attribute VB_Description = ""Desc""
'@ModuleAttribute VB_Description, ""Desc""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new AttributeValueOutOfSyncInspection(state), CodeKind.AttributesCode);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void VbExtKeyModuleAttributeOutOfSync_QuickFixWorks()
        {
            const string inputCode =
                @"Attribute VB_Ext_Key = ""Key"", ""NotValue""
Attribute VB_Ext_Key = ""OtherKey"", ""OtherValue""
'@ModuleAttribute VB_Ext_Key, ""Key"", ""Value""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            const string expectedCode =
                @"Attribute VB_Ext_Key = ""Key"", ""Value""
Attribute VB_Ext_Key = ""OtherKey"", ""OtherValue""
'@ModuleAttribute VB_Ext_Key, ""Key"", ""Value""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new AttributeValueOutOfSyncInspection(state), CodeKind.AttributesCode);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void MemberAttributeOutOfSync_QuickFixWorks()
        {
            const string inputCode =
                @"'@MemberAttribute VB_Description, ""Desc""
Public Sub Foo()
Attribute Foo.VB_Description = ""NotDesc""
    Const const1 As Integer = 9
End Sub";

            const string expectedCode =
                @"'@MemberAttribute VB_Description, ""Desc""
Public Sub Foo()
Attribute Foo.VB_Description = ""Desc""
    Const const1 As Integer = 9
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new AttributeValueOutOfSyncInspection(state), CodeKind.AttributesCode);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void VbExtKeyMemberAttributeOutOfSync_QuickFixWorks()
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
Attribute Foo.VB_Ext_Key = ""Key"", ""Value""
Attribute Foo.VB_Ext_Key = ""OtherKey"", ""OtherValue""
    Const const1 As Integer = 9
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new AttributeValueOutOfSyncInspection(state), CodeKind.AttributesCode);
            Assert.AreEqual(expectedCode, actualCode);
        }

        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new AdjustAttributeValuesQuickFix(new AttributesUpdater(state));
        }
    }
}