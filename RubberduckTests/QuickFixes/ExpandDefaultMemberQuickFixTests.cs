using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class ExpandDefaultMemberQuickFixTests : QuickFixTestBase
    {
        [Test]
        [Category("QuickFixes")]
        public void DictionaryAccessExpression_QuickFixWorks()
        {
            var class1Code = @"
Public Sub Foo()
Attribute Foo.VB_UserMemId = 0
End Sub
";

            var class2Code = @"
Public Function Baz() As Class1
Attribute Baz.VB_UserMemId = 0
    Set Baz = New Class1
End Function
";

            var moduleCode = $@"
Private Function Foo() As Variant 
    Dim cls As New Class2
    cls
End Function
";

            var expectedModuleCode = $@"
Private Function Foo() As Variant 
    Dim cls As New Class2
    cls.Baz
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var actualModuleCode = ApplyQuickFixToFirstInspectionResult(vbe.Object, "Module1", state => new ObjectWhereProcedureIsRequiredInspection(state), CodeKind.AttributesCode);
            Assert.AreEqual(expectedModuleCode, actualModuleCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void NonParameterizedProcedureCoercion_ExplicitCall_QuickFixWorks()
        {
            var class1Code = @"
Public Sub Foo()
Attribute Foo.VB_UserMemId = 0
End Sub
";

            var class2Code = @"
Public Function Baz() As Class1
Attribute Baz.VB_UserMemId = 0
    Set Baz = New Class1
End Function
";

            var moduleCode = $@"
Private Function Foo() As Variant 
    Dim cls As New Class2
    Call cls
End Function
";

            var expectedModuleCode = $@"
Private Function Foo() As Variant 
    Dim cls As New Class2
    Call cls.Baz
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var actualModuleCode = ApplyQuickFixToFirstInspectionResult(vbe.Object, "Module1", state => new ObjectWhereProcedureIsRequiredInspection(state));
            Assert.AreEqual(expectedModuleCode, actualModuleCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void NonParameterizedProcedureCoercionDefaultMemberAccessOnDefaultMemberArrayAccess_QuickFixWorks()
        {
            var class1Code = @"
Public Sub Foo()
Attribute Foo.VB_UserMemId = 0
End Sub
";

            var class2Code = @"
Public Function Baz() As Class1()
Attribute Baz.VB_UserMemId = 0
    Set Baz = New Class1
End Function
";

            var moduleCode = $@"
Private Function Foo() As Variant 
    Dim cls As New Class2
    cls(42)
End Function
";

            var expectedModuleCode = $@"
Private Function Foo() As Variant 
    Dim cls As New Class2
    cls(42).Foo
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var actualModuleCode = ApplyQuickFixToFirstInspectionResult(vbe.Object, "Module1", state => new ObjectWhereProcedureIsRequiredInspection(state));
            Assert.AreEqual(expectedModuleCode, actualModuleCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void NonParameterizedProcedureCoercionDefaultMemberAccessOnDefaultMemberArrayAccess_ExplicitCall_QuickFixWorks()
        {
            var class1Code = @"
Public Sub Foo()
Attribute Foo.VB_UserMemId = 0
End Sub
";

            var class2Code = @"
Public Function Baz() As Class1()
Attribute Baz.VB_UserMemId = 0
    Set Baz = New Class1
End Function
";

            var moduleCode = $@"
Private Function Foo() As Variant 
    Dim cls As New Class2
    Call cls(42)
End Function
";

            var expectedModuleCode = $@"
Private Function Foo() As Variant 
    Dim cls As New Class2
    Call cls(42).Foo
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var actualModuleCode = ApplyQuickFixToFirstInspectionResult(vbe.Object, "Module1", state => new ObjectWhereProcedureIsRequiredInspection(state));
            Assert.AreEqual(expectedModuleCode, actualModuleCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void OrdinaryIndexedDefaultMemberAccessExpression_QuickFixWorks()
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
    Set Foo = cls(""newClassObject"")
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

            var actualModuleCode = ApplyQuickFixToAllInspectionResults(vbe.Object, "Module1", state => new IndexedDefaultMemberAccessInspection(state));
            Assert.AreEqual(expectedModuleCode, actualModuleCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void RecursiveIndexedDefaultMemberAccessExpression_QuickFixWorks()
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
    Set Foo = cls(""newClassObject"")
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

            var actualModuleCode = ApplyQuickFixToAllInspectionResults(vbe.Object, "Module1", state => new IndexedRecursiveDefaultMemberAccessInspection(state));
            Assert.AreEqual(expectedModuleCode, actualModuleCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void DoubleOrdinaryIndexedDefaultMemberAccessExpression_QuickFixWorks()
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
    Set Foo = cls(""newClassObject"")(""whatever"")
End Function
";

            var expectedModuleCode = @"
Private Function Foo() As Class2 
    Dim cls As New Class1
    Set Foo = cls.Foo(""newClassObject"").Baz(""whatever"")
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var actualModuleCode = ApplyQuickFixToAllInspectionResults(vbe.Object, "Module1", state => new IndexedDefaultMemberAccessInspection(state));
            Assert.AreEqual(expectedModuleCode, actualModuleCode);
        }

        [Test]
        [Category("Inspections")]
        public void DoubleRecursiveIndexedDefaultMemberAccessExpression_QuickFixWorks()
        {
            var class1Code = @"
Public Function Foo() As Class2
Attribute Foo.VB_UserMemId = 0
End Function
";

            var class2Code = @"
Public Function Baz(bar As String) As Class1
Attribute Baz.VB_UserMemId = 0
End Function
";

            var moduleCode = @"
Private Function Foo() As Class2 
    Dim cls As New Class1
    Set Foo = cls(""newClassObject"")(""whatever"")
End Function
";

            var expectedModuleCode = @"
Private Function Foo() As Class2 
    Dim cls As New Class1
    Set Foo = cls.Foo().Baz(""newClassObject"").Foo().Baz(""whatever"")
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var actualModuleCode = ApplyQuickFixToAllInspectionResults(vbe.Object, "Module1", state => new IndexedRecursiveDefaultMemberAccessInspection(state));
            Assert.AreEqual(expectedModuleCode, actualModuleCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void OrdinaryIndexedDefaultMemberAccessOnDefaultMemberArrayAccess_QuickFixWorks()
        {
            var class1Code = @"
Public Function Foo() As Class2()
Attribute Foo.VB_UserMemId = 0
End Function
";

            var class2Code = @"
Public Function Baz(bar As String) As Class2
Attribute Baz.VB_UserMemId = 0
End Function
";

            var class3Code = @"
Public Function Baz() As Class2
Attribute Baz.VB_UserMemId = 0
End Function
";

            var moduleCode = @"
Private Function Foo() As Class2 
    Dim cls As New Class1
    Set Foo = cls(0)(""newClassObject"")
End Function
";

            var expectedModuleCode = @"
Private Function Foo() As Class2 
    Dim cls As New Class1
    Set Foo = cls(0).Baz(""newClassObject"")
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Class3", class3Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var actualModuleCode = ApplyQuickFixToAllInspectionResults(vbe.Object, "Module1", state => new IndexedDefaultMemberAccessInspection(state));
            Assert.AreEqual(expectedModuleCode, actualModuleCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void RecursiveIndexedDefaultMemberAccessOnDefaultMemberArrayAccess_QuickFixWorks()
        {
            var class1Code = @"
Public Function Foo() As Class3()
Attribute Foo.VB_UserMemId = 0
End Function
";

            var class2Code = @"
Public Function Baz(bar As String) As Class2
Attribute Baz.VB_UserMemId = 0
End Function
";

            var class3Code = @"
Public Function Bar() As Class2
Attribute Bar.VB_UserMemId = 0
End Function
";

            var moduleCode = @"
Private Function Foo() As Class2 
    Dim cls As New Class1
    Set Foo = cls(0)(""newClassObject"")
End Function
";

            var expectedModuleCode = @"
Private Function Foo() As Class2 
    Dim cls As New Class1
    Set Foo = cls(0).Bar().Baz(""newClassObject"")
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Class3", class3Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var actualModuleCode = ApplyQuickFixToAllInspectionResults(vbe.Object, "Module1", state => new IndexedRecursiveDefaultMemberAccessInspection(state));
            Assert.AreEqual(expectedModuleCode, actualModuleCode);
        }

        [Test]
        [Category("QuickFixes")]
        [TestCase("Foo = cls", "Foo = cls.Foo")]
        [TestCase("cls = bar", "cls.Foo = bar")]
        public void OrdinaryImplicitDefaultMemberAccessExpression_QuickFixWorks(string statement, string resultStatement)
        {
            var class1Code = @"
Public Function Foo() As Long
Attribute Foo.VB_UserMemId = 0
End Function
";

            var moduleCode = $@"
Private Function Foo() As Long 
    Dim cls As New Class1
    Dim bar As Long
    {statement}
End Function
";

            var expectedModuleCode = $@"
Private Function Foo() As Long 
    Dim cls As New Class1
    Dim bar As Long
    {resultStatement}
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var actualModuleCode = ApplyQuickFixToAllInspectionResults(vbe.Object, "Module1", state => new ImplicitDefaultMemberAccessInspection(state));
            Assert.AreEqual(expectedModuleCode, actualModuleCode);
        }

        [Test]
        [Category("QuickFixes")]
        [TestCase("Foo = cls", "Foo = cls.Foo().Baz")]
        [TestCase("cls = bar", "cls.Foo().Baz = bar")]
        public void RecursiveImplicitDefaultMemberAccessExpression_QuickFixWorks(string statement, string resultStatement)
        {
            var class1Code = @"
Public Function Foo() As Class2
Attribute Foo.VB_UserMemId = 0
End Function
";

            var class2Code = @"
Public Function Baz() As Long
Attribute Baz.VB_UserMemId = 0
End Function
";

            var moduleCode = $@"
Private Function Foo() As Long 
    Dim cls As New Class1
    Dim bar As Long
    {statement}
End Function
";

            var expectedModuleCode = $@"
Private Function Foo() As Long 
    Dim cls As New Class1
    Dim bar As Long
    {resultStatement}
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var actualModuleCode = ApplyQuickFixToAllInspectionResults(vbe.Object, "Module1", state => new ImplicitRecursiveDefaultMemberAccessInspection(state));
            Assert.AreEqual(expectedModuleCode, actualModuleCode);
        }

        [Test]
        [Category("QuickFixes")]
        [TestCase("Foo = cls(0)", "Foo = cls.Foo(0).Baz")]
        [TestCase("cls(0) = bar", "cls.Foo(0).Baz = bar")]
        public void OrdinaryImplicitDefaultMemberAccessOnDefaultMemberArrayAccess_QuickFixWorks(string statement, string resultStatement)
        {
            var class1Code = @"
Public Function Foo() As Class2()
Attribute Foo.VB_UserMemId = 0
End Function
";

            var class2Code = @"
Public Function Baz() As Long
Attribute Baz.VB_UserMemId = 0
End Function
";

            var moduleCode = $@"
Private Function Foo() As Long 
    Dim cls As New Class1
    Dim bar As Long
    {statement}
End Function
";

            var expectedModuleCode = $@"
Private Function Foo() As Long 
    Dim cls As New Class1
    Dim bar As Long
    {resultStatement}
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var actualModuleCode = ApplyQuickFixToAllInspectionResults(vbe.Object, "Module1", state => new ImplicitDefaultMemberAccessInspection(state));
            Assert.AreEqual(expectedModuleCode, actualModuleCode);
        }

        [Test]
        [Category("QuickFixes")]
        [TestCase("Foo = cls(0)", "Foo = cls.Barz().Foo(0).Bar().Baz")]
        [TestCase("cls(0) = bar", "cls.Barz().Foo(0).Bar().Baz = bar")]
        public void RecursiveImplicitDefaultMemberAccessOnRecursiveDefaultMemberArrayAccess_QuickFixWorks(string statement, string resultStatement)
        {
            var class1Code = @"
Public Function Foo() As Class4()
Attribute Foo.VB_UserMemId = 0
End Function
";

            var class2Code = @"
Public Function Baz() As Long
Attribute Baz.VB_UserMemId = 0
End Function
";

            var class3Code = @"
Public Function Barz() As Class1
Attribute Barz.VB_UserMemId = 0
End Function
";

            var class4Code = @"
Public Function Bar() As Class2
Attribute Bar.VB_UserMemId = 0
End Function
";

            var moduleCode = $@"
Private Function Foo() As Long 
    Dim cls As New Class3
    Dim bar As Long
    {statement}
End Function
";

            var expectedModuleCode = $@"
Private Function Foo() As Long 
    Dim cls As New Class3
    Dim bar As Long
    {resultStatement}
End Function
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Class3", class3Code, ComponentType.ClassModule),
                ("Class4", class4Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var actualModuleCode = ApplyQuickFixToAllInspectionResults(vbe.Object, "Module1", state => new ImplicitRecursiveDefaultMemberAccessInspection(state));
            Assert.AreEqual(expectedModuleCode, actualModuleCode);
        }

        [Test]
        [Category("Inspections")]
        public void BothSidesOfAssignmentHaveDefaultMemberAccess_NoExplicitLet_QuickFixWorks()
        {
            var class1Code = @"
Public Function Foo() As Long
Attribute Foo.VB_UserMemId = 0
End Function
";

            var class2Code = @"
Public Property Let Baz(RHS As Long)
Attribute Baz.VB_UserMemId = 0
End Property
";

            var moduleCode = $@"
Private Sub Bar() 
    Dim cls1 As Class1
    Dim cls2 As Class2
    Set cls1 = New Class1
    Set cls2 = New Class2
    cls2 = cls1
End Sub
";

            var expectedModuleCode = $@"
Private Sub Bar() 
    Dim cls1 As Class1
    Dim cls2 As Class2
    Set cls1 = New Class1
    Set cls2 = New Class2
    cls2.Baz = cls1.Foo
End Sub
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Class2", class2Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var actualModuleCode = ApplyQuickFixToAllInspectionResults(vbe.Object, "Module1", state => new SuspiciousLetAssignmentInspection(state));
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
    bar = wkb.Sheets(""MySheet"").Range(""A1"").Value
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

            var actualModuleCode = ApplyQuickFixToFirstInspectionResult(vbe.Object, "Module1", state => new IndexedDefaultMemberAccessInspection(state));
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
    Set bar = baz.Columns(""A"")
    Set bar = baz.Cells(1,1)
End Sub
";

            var expectedModuleCode = @"
Private Sub Foo() 
    Dim baz As Excel.Range
    Dim bar As Variant
    Set bar = baz.Columns.Item(""A"")
    Set bar = baz.Cells.Item(1,1)
End Sub
";

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, moduleCode)
                .AddReference(ReferenceLibrary.Excel)
                .AddProjectToVbeBuilder()
                .Build();

            var actualModuleCode = ApplyQuickFixToAllInspectionResults(vbe.Object, "Module1", state => new IndexedDefaultMemberAccessInspection(state));
            Assert.AreEqual(expectedModuleCode, actualModuleCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void SpecialDefaultMembersAreReplacedBasedOnNameAndArgumentNumber()
        {
            var moduleCode = @"
Private Sub Foo() 
    Dim baz As Excel.Range
    Dim bar As Variant
    bar = baz
End Sub
";

            var expectedModuleCode = @"
Private Sub Foo() 
    Dim baz As Excel.Range
    Dim bar As Variant
    bar = baz.Value
End Sub
";

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, moduleCode)
                .AddReference(ReferenceLibrary.Excel)
                .AddProjectToVbeBuilder()
                .Build();

            var actualModuleCode = ApplyQuickFixToFirstInspectionResult(vbe.Object, "Module1", state => new ImplicitDefaultMemberAccessInspection(state));
            Assert.AreEqual(expectedModuleCode, actualModuleCode);
        }

        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new ExpandDefaultMemberQuickFix(state);
        }
    }
}