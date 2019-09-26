using NUnit.Framework;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
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

            var actualModuleCode = ApplyQuickFixToFirstInspectionResult(vbe.Object, "Module1", state => new ObjectWhereProcedureIsRequiredInspection(state), CodeKind.AttributesCode);
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

            var actualModuleCode = ApplyQuickFixToFirstInspectionResult(vbe.Object, "Module1", state => new ObjectWhereProcedureIsRequiredInspection(state), CodeKind.AttributesCode);
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

            var actualModuleCode = ApplyQuickFixToFirstInspectionResult(vbe.Object, "Module1", state => new ObjectWhereProcedureIsRequiredInspection(state), CodeKind.AttributesCode);
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

            var actualModuleCode = ApplyQuickFixToAllInspectionResults(vbe.Object, "Module1", state => new IndexedDefaultMemberAccessInspection(state), CodeKind.AttributesCode);
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

            var actualModuleCode = ApplyQuickFixToAllInspectionResults(vbe.Object, "Module1", state => new IndexedRecursiveDefaultMemberAccessInspection(state), CodeKind.AttributesCode);
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

            var actualModuleCode = ApplyQuickFixToAllInspectionResults(vbe.Object, "Module1", state => new IndexedDefaultMemberAccessInspection(state), CodeKind.AttributesCode);
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

            var actualModuleCode = ApplyQuickFixToAllInspectionResults(vbe.Object, "Module1", state => new IndexedRecursiveDefaultMemberAccessInspection(state), CodeKind.AttributesCode);
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

            var actualModuleCode = ApplyQuickFixToAllInspectionResults(vbe.Object, "Module1", state => new IndexedDefaultMemberAccessInspection(state), CodeKind.AttributesCode);
            Assert.AreEqual(expectedModuleCode, actualModuleCode);
        }

        [Test]
        [Category("Inspections")]
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

            var actualModuleCode = ApplyQuickFixToAllInspectionResults(vbe.Object, "Module1", state => new IndexedRecursiveDefaultMemberAccessInspection(state), CodeKind.AttributesCode);
            Assert.AreEqual(expectedModuleCode, actualModuleCode);
        }

        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new ExpandDefaultMemberQuickFix(state);
        }
    }
}