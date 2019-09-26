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

        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new ExpandDefaultMemberQuickFix(state);
        }
    }
}