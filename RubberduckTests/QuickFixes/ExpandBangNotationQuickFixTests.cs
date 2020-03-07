using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    public class ExpandBangNotationQuickFixTests : QuickFixTestBase
    {
        [Test]
        [Category("QuickFixes")]
        public void DictionaryAccessExpression_QuickFixWorks()
        {
            var class1Code = @"
Public Function Foo(bar As String) As Class2
Attribute Foo.VB_UserMemId = 0
End Function
";

            var class2Code = @"
Public Function Baz(bar As String) As Class2
Attribute Baz.VB_UserMemId = 0
End Function
";

            var moduleCode = @"
Private Function Foo() As Class2 
    Dim cls As New Class1
    Set Foo = cls!newClassObject
End Function
";

            var expectedModuleCode = @"
Private Function Foo() As Class2 
    Dim cls As New Class1
    Set Foo = cls.Foo(""newClassObject"")
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var actualModuleCode = ApplyQuickFixToFirstInspectionResult(vbe.Object, "Module1", state => new UseOfBangNotationInspection(state));
            Assert.AreEqual(expectedModuleCode, actualModuleCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void RecursiveDictionaryAccessExpression_QuickFixWorks()
        {
            var class1Code = @"
Public Function Foo() As Class2
Attribute Foo.VB_UserMemId = 0
End Function
";

            var class2Code = @"
Public Function Baz(bar As String) As Class2
Attribute Baz.VB_UserMemId = 0
End Function
";

            var moduleCode = @"
Private Function Foo() As Class2 
    Dim cls As New Class1
    Set Foo = cls!newClassObject
End Function
";

            var expectedModuleCode = @"
Private Function Foo() As Class2 
    Dim cls As New Class1
    Set Foo = cls.Foo().Baz(""newClassObject"")
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var actualModuleCode = ApplyQuickFixToFirstInspectionResult(vbe.Object, "Module1", state => new UseOfRecursiveBangNotationInspection(state));
            Assert.AreEqual(expectedModuleCode, actualModuleCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void WithDictionaryAccessExpression_QuickFixWorks()
        {
            var class1Code = @"
Public Function Foo(bar As String) As Class2
Attribute Foo.VB_UserMemId = 0
End Function
";

            var class2Code = @"
Public Function Baz(bar As String) As Class2
Attribute Baz.VB_UserMemId = 0
End Function
";

            var moduleCode = @"
Private Function Foo() As Class2 
    With New Class1
        Set Foo = !newClassObject
    End With
End Function
";

            var expectedModuleCode = @"
Private Function Foo() As Class2 
    With New Class1
        Set Foo = .Foo(""newClassObject"")
    End With
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var actualModuleCode = ApplyQuickFixToFirstInspectionResult(vbe.Object, "Module1", state => new UseOfBangNotationInspection(state));
            Assert.AreEqual(expectedModuleCode, actualModuleCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void RecursiveWithDictionaryAccessExpression_QuickFixWorks()
        {
            var class1Code = @"
Public Function Foo() As Class2
Attribute Foo.VB_UserMemId = 0
End Function
";

            var class2Code = @"
Public Function Baz(bar As String) As Class2
Attribute Baz.VB_UserMemId = 0
End Function
";

            var moduleCode = @"
Private Function Foo() As Class2 
    With New Class1
        Set Foo = !newClassObject
    End With
End Function
";

            var expectedModuleCode = @"
Private Function Foo() As Class2 
    With New Class1
        Set Foo = .Foo().Baz(""newClassObject"")
    End With
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var actualModuleCode = ApplyQuickFixToFirstInspectionResult(vbe.Object, "Module1", state => new UseOfRecursiveBangNotationInspection(state));
            Assert.AreEqual(expectedModuleCode, actualModuleCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void OtherwiseIllegalDeclarationNamesAreEnclosedInBrackets()
        {
            var moduleCode = @"
Private Sub Foo() 
    Dim wkb As Excel.Workbook
    Dim bar As Variant
    bar = wkb.Sheets!MySheet.Range(""A1"").Value
End Sub
";

            var expectedModuleCode = @"
Private Sub Foo() 
    Dim wkb As Excel.Workbook
    Dim bar As Variant
    bar = wkb.Sheets.[_Default](""MySheet"").Range(""A1"").Value
End Sub
";

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, moduleCode)
                .AddReference(ReferenceLibrary.Excel)
                .AddProjectToVbeBuilder()
                .Build();

            var actualModuleCode = ApplyQuickFixToFirstInspectionResult(vbe.Object, "Module1", state => new UseOfBangNotationInspection(state));
            Assert.AreEqual(expectedModuleCode, actualModuleCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void SpecialDefaultMembersAreReplacedBasedOnName()
        {
            var moduleCode = @"
Private Sub Foo() 
    Dim baz As Excel.Range
    Dim bar As Variant
    Set bar = baz.Columns!A
End Sub
";

            var expectedModuleCode = @"
Private Sub Foo() 
    Dim baz As Excel.Range
    Dim bar As Variant
    Set bar = baz.Columns.Item(""A"")
End Sub
";

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, moduleCode)
                .AddReference(ReferenceLibrary.Excel)
                .AddProjectToVbeBuilder()
                .Build();

            var actualModuleCode = ApplyQuickFixToFirstInspectionResult(vbe.Object, "Module1", state => new UseOfBangNotationInspection(state));
            Assert.AreEqual(expectedModuleCode, actualModuleCode);
        }

        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new ExpandBangNotationQuickFix(state);
        }
    }
}