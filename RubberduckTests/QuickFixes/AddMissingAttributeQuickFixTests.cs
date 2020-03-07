using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
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
        //See issue #5268 at https://github.com/rubberduck-vba/Rubberduck/issues/5268
        public void MissingMemberAttribute_ExcelHotkey_QuickFixWorks()
        {
            const string inputCode =
                @"'@ExcelHotkey ""T""
Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            const string expectedCode =
                @"'@ExcelHotkey ""T""
Public Sub Foo()
Attribute Foo.VB_ProcData.VB_Invoke_Func = ""T\n14""
    Const const1 As Integer = 9
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new MissingAttributeInspection(state), CodeKind.AttributesCode);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void MissingMemberAttributeOnConditionalCompilation_QuickFixWorks()
        {
            const string inputCode =
                @"'@Description(""Desc"")
#If False Then
    Private Sub Bar(ByVal arg As Long)
#Else
    Private Sub Foo(ByVal arg As Long)
#End If
End Sub";

            const string expectedCode =
                    @"'@Description(""Desc"")
#If False Then
    Private Sub Bar(ByVal arg As Long)
#Else
    Private Sub Foo(ByVal arg As Long)
Attribute Foo.VB_Description = ""Desc""
#End If
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new MissingAttributeInspection(state), CodeKind.AttributesCode);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void MissingMemberAttributeOnDeclareStatement_QuickFixWorks()
        {
            const string inputCode =
                @"'@Description(""Desc"")
Private Declare Sub CopyMemory Lib ""kernel32.dll"" Alias ""RtlMoveMemory"" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)
";

            const string expectedCode =
                @"'@Description(""Desc"")
Private Declare Sub CopyMemory Lib ""kernel32.dll"" Alias ""RtlMoveMemory"" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)
Attribute CopyMemory.VB_Description = ""Desc""
";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new MissingAttributeInspection(state), CodeKind.AttributesCode);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void MissingMemberAttributeOnConditionalCompilation_DeclareStatement_QuickFixWorks()
        {
            const string inputCode =
                @"'@Description(""Desc"")
#If False Then
    Private Declare PtrSafe Sub CopyMemory Lib ""kernel32.dll"" Alias ""RtlMoveMemory"" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)
#Else
    Private Declare Sub CopyMemory Lib ""kernel32.dll"" Alias ""RtlMoveMemory""(ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)
#End If";

            const string expectedCode =
                @"'@Description(""Desc"")
#If False Then
    Private Declare PtrSafe Sub CopyMemory Lib ""kernel32.dll"" Alias ""RtlMoveMemory"" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)
#Else
    Private Declare Sub CopyMemory Lib ""kernel32.dll"" Alias ""RtlMoveMemory""(ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)
Attribute CopyMemory.VB_Description = ""Desc""
#End If";

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