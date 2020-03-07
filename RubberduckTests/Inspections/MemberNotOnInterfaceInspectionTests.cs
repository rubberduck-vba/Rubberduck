using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class MemberNotOnInterfaceInspectionTests : InspectionTestsBase
    {
        private int ArrangeParserAndGetResultCount(string inputCode, ReferenceLibrary library = ReferenceLibrary.Scripting)
            => InspectionResultsForModules(("Codez", inputCode, ComponentType.StandardModule), library).Count();

        [Test]
        [Category("Inspections")]
        public void MemberNotOnInterface_ReturnsResult_UnDeclaredMember()
        {
            const string inputCode =
                @"Sub Foo()
    Dim dict As Dictionary
    Set dict = New Dictionary
    dict.NonMember
End Sub";
            Assert.AreEqual(1, ArrangeParserAndGetResultCount(inputCode));
        }

        [Test]
        [Category("Inspections")]
        public void MemberNotOnInterface_ReturnsResult_UnDeclaredInterfaceMember()
        {
            const string inputCode =
                @"Sub Foo()
    Dim dict As Dictionary
    Set dict = New Dictionary
    dict.NonMember
End Sub";
            Assert.AreEqual(1, ArrangeParserAndGetResultCount(inputCode));
        }

        [Test]
        [Category("Inspections")]
        public void MemberNotOnInterface_ReturnsResult_ApplicationObject()
        {
            const string inputCode =
                @"Sub Foo()
    Application.NonMember
End Sub";
            Assert.AreEqual(1, ArrangeParserAndGetResultCount(inputCode, ReferenceLibrary.Excel));
        }

        [Test]
        [Category("Inspections")]
        public void MemberNotOnInterface_ReturnsResult_UnDeclaredMemberOnParameter()
        {
            const string inputCode =
                @"Sub Foo(dict As Dictionary)
    dict.NonMember
End Sub";
            Assert.AreEqual(1, ArrangeParserAndGetResultCount(inputCode));
        }

        [Test]
        [Category("Inspections")]
        public void MemberNotOnInterface_DoesNotReturnResult_DeclaredMember()
        {
            const string inputCode =
                @"Sub Foo()
    Dim dict As Dictionary
    Set dict = New Dictionary
    Debug.Print dict.Count
End Sub";
            Assert.AreEqual(0, ArrangeParserAndGetResultCount(inputCode));
        }

        [Test]
        [Category("Inspections")]
        public void MemberNotOnInterface_DoesNotReturnResult_NonExtensible()
        {
            const string inputCode =
                @"Sub Foo()
    Dim x As File
    Debug.Print x.NonMember
End Sub";
            Assert.AreEqual(0, ArrangeParserAndGetResultCount(inputCode));
        }

        [Test]
        [Category("Inspections")]
        public void MemberNotOnInterface_ReturnsResult_WithBlock()
        {
            const string inputCode =
                @"Sub Foo()
    Dim dict As New Dictionary
    With dict
        .NonMember
    End With
End Sub";
            Assert.AreEqual(1, ArrangeParserAndGetResultCount(inputCode));
        }

        [Test]
        [Category("Inspections")]
        public void MemberNotOnInterface_DoesNotReturnResult_BangNotation()
        {
            const string inputCode =
                @"Sub Foo()
    Dim dict As Dictionary
    Set dict = New Dictionary
    dict!SomeIdentifier = 42
End Sub";
            Assert.AreEqual(0, ArrangeParserAndGetResultCount(inputCode));
        }

        [Test]
        [Category("Inspections")]
        public void MemberNotOnInterface_DoesNotReturnResult_WithBlockBangNotation()
        {
            const string inputCode =
                @"Sub Foo()
    Dim dict As New Dictionary
    With dict
        !SomeIdentifier = 42
    End With
End Sub";
            Assert.AreEqual(0, ArrangeParserAndGetResultCount(inputCode));
        }

        [Test]
        [Category("Inspections")]
        public void MemberNotOnInterface_DoesNotReturnResult_ProjectReference()
        {
            const string inputCode =
                @"Sub Foo()
    Dim dict As Scripting.Dictionary
End Sub";
            Assert.AreEqual(0, ArrangeParserAndGetResultCount(inputCode));
        }

        [Test]
        [Category("Inspections")]
        public void MemberNotOnInterface_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"Sub Foo(dict As Dictionary)
    Dim dict As Dictionary
    Set dict = New Dictionary
    '@Ignore MemberNotOnInterface
    dict.NonMember
End Sub";
            Assert.AreEqual(0, ArrangeParserAndGetResultCount(inputCode));
        }

        [Test]
        [Category("Inspections")]
        public void MemberNotOnInterface_WithNewReturnsResult()
        {
            const string inputCode =
                @"Sub Foo()
    With New Dictionary
        .FooBar
    End With
End Sub";
            Assert.AreEqual(1, ArrangeParserAndGetResultCount(inputCode));
        }

        [Test]
        [Category("Inspections")]
        //See https://github.com/rubberduck-vba/Rubberduck/issues/4308 
        public void MemberNotOnInterface_ProcedureArgument()
        {
            const string inputCode =
                @"Sub Foo()
    Dim fooBaz As Dictionary
    Set fooBaz = New Dictionary 
    Bar fooBaz.FooBar
End Sub

Private Sub Bar(baz As Long)
End Sub";
            Assert.AreEqual(1, ArrangeParserAndGetResultCount(inputCode));
        }

        [Test]
        [Category("Inspections")]
        public void MemberNotOnInterface_FunctionArgument()
        {
            const string inputCode =
                @"Sub Foo()
    Dim fooBaz As Dictionary
    Dim fooBar As Variant 
    Set fooBaz = New Dictionary 
    fooBar = Bar(fooBaz.FooBar)
End Sub

Private Function Bar(baz As Long) As Variant
End Function";
            Assert.AreEqual(1, ArrangeParserAndGetResultCount(inputCode));
        }

        [Test]
        [Category("Inspections")]
        public void MemberNotOnInterface_Expression()
        {
            const string inputCode =
                @"Sub Foo()
    Dim fooBaz As Dictionary
    Dim fooBar As Variant 
    Set fooBaz = New Dictionary 
    fooBar = 1 + fooBaz.FooBar
End Sub

Private Function Bar(baz As Long) As Variant
End Function";
            Assert.AreEqual(1, ArrangeParserAndGetResultCount(inputCode));
        }

        [Test]
        [Category("Inspections")]
        public void MemberNotOnInterface_Expression_BothSides()
        {
            const string inputCode =
                @"Sub Foo()
    Dim fooBaz As Dictionary
    Dim fooBar As Variant 
    Set fooBaz = New Dictionary 
    fooBar = fooBaz.NotThere + fooBaz.FooBar
End Sub

Private Function Bar(baz As Long) As Variant
End Function";
            Assert.AreEqual(2, ArrangeParserAndGetResultCount(inputCode));
        }

        [Test]
        [Category("Inspections")]
        public void MemberNotOnInterface_DeepExpression()
        {
            const string inputCode =
                @"Sub Foo()
    Dim fooBaz As Dictionary
    Dim fooBar As Variant 
    Set fooBaz = New Dictionary 
    fooBar = 1 + (1 + (1 + (1 + fooBaz.FooBar)))
End Sub

Private Function Bar(baz As Long) As Variant
End Function";
            Assert.AreEqual(1, ArrangeParserAndGetResultCount(inputCode));
        }

        [Test]
        [Category("Inspections")]
        public void MemberNotOnInterface_ExpressionInFunctionArgument()
        {
            const string inputCode =
                @"Sub Foo()
    Dim fooBaz As Dictionary
    Dim fooBar As Variant 
    Set fooBaz = New Dictionary 
    fooBar = Bar(1 + fooBaz.FooBar)
End Sub

Private Function Bar(baz As Long) As Variant
End Function";
            Assert.AreEqual(1, ArrangeParserAndGetResultCount(inputCode));
        }

        [Test]
        [Category("Inspections")]
        public void MemberNotOnInterface_FunctionArgumentInFunctionArgument()
        {
            const string inputCode =
                @"Sub Foo()
    Dim fooBaz As Dictionary
    Dim fooBar As Variant 
    Set fooBaz = New Dictionary 
    fooBar = Bar(Bar(fooBaz.FooBar))
End Sub

Private Function Bar(baz As Long) As Variant
End Function";
            Assert.AreEqual(1, ArrangeParserAndGetResultCount(inputCode));
        }

        [Test]
        [Category("Inspections")]
        public void MemberNotOnInterface_FunctionArgumentInExpressionInFunctionArgument()
        {
            const string inputCode =
                @"Sub Foo()
    Dim fooBaz As Dictionary
    Dim fooBar As Variant 
    Set fooBaz = New Dictionary 
    fooBar = Bar(1 + Bar(fooBaz.FooBar))
End Sub

Private Function Bar(baz As Long) As Variant
End Function";
            Assert.AreEqual(1, ArrangeParserAndGetResultCount(inputCode));
        }

        [Test]
        [Category("Inspections")]
        public void MemberNotOnInterface_FunctionArgumentInExpressionInProcedureArgument()
        {
            const string inputCode =
                @"Sub Foo()
    Dim fooBaz As Dictionary
    Dim fooBar As Variant 
    Set fooBaz = New Dictionary 
    Barr 1 + Bar(fooBaz.FooBar)
End Sub

Private Function Bar(baz As Long) As Variant
End Function

Private Sub Barr(baz As Long)
End Sub";
            Assert.AreEqual(1, ArrangeParserAndGetResultCount(inputCode));
        }

        [Test]
        [Category("Inspections")]
        public void MemberNotOnInterface_FunctionArgumentInExpressionInProcedureArgument_ExplicitCall()
        {
            const string inputCode =
                @"Sub Foo()
    Dim fooBaz As Dictionary
    Dim fooBar As Variant 
    Set fooBaz = New Dictionary 
    Call Barr(1 + Bar(fooBaz.FooBar))
End Sub

Private Function Bar(baz As Long) As Variant
End Function

Private Sub Barr(baz As Long)
End Sub";
            Assert.AreEqual(1, ArrangeParserAndGetResultCount(inputCode));
        }

        [Test]
        [Category("Inspections")]
        public void MemberNotOnInterface_InOutputList()
        {
            const string inputCode =
                @"Sub Foo()
    Dim fooBaz As Dictionary
    Dim fooBar As Variant 
    Set fooBaz = New Dictionary 
    Debug.Print fooBaz.FooBar; Spc(fooBaz.NotThere); Tab(fooBaz.Neither)
End Sub

Private Function Bar(baz As Long) As Variant
End Function

Private Sub Barr(baz As Long)
End Sub";
            Assert.AreEqual(3, ArrangeParserAndGetResultCount(inputCode));
        }

        [Test]
        [Category("Inspections")]
        public void MemberNotOnInterface_FunctionArgumentInExpressionInOutputList()
        {
            const string inputCode =
                @"Sub Foo()
    Dim fooBaz As Dictionary
    Dim fooBar As Variant 
    Set fooBaz = New Dictionary 
    Debug.Print 1 + Bar(fooBaz.FooBar)
End Sub

Private Function Bar(baz As Long) As Variant
End Function

Private Sub Barr(baz As Long)
End Sub";
            Assert.AreEqual(1, ArrangeParserAndGetResultCount(inputCode));
        }

        [Test]
        [Category("Inspections")]
        public void MemberNotOnInterface_DoesNotReturnResult_WithNewBlockBangNotation()
        {
            const string inputCode =
                @"Sub Foo()
    With New Dictionary
        !FooBar = 42
    End With
End Sub";
            Assert.AreEqual(0, ArrangeParserAndGetResultCount(inputCode));
        }

        [Test]
        [Category("Inspections")]
        public void MemberNotOnInterface_DoesNotReturnResult_WithNewBlockOnInterface()
        {
            const string inputCode =
                @"Sub Foo()
    With New Dictionary
        .Add 42, 42
    End With
End Sub";
            Assert.AreEqual(0, ArrangeParserAndGetResultCount(inputCode));
        }

        [Test]
        [Category("Inspections")]
        public void MemberNotOnInterface_CatchesInvalidUseOfMember()
        {
            const string userForm1Code = @"
Private mfooBar As String

Public Property Let FooBar(value As String)
    mfooBar = value
End Property

Public Property Get FooBar() As String
    FooBar = mfooBar
End Property
";

            const string analyzedCode = @"Option Explicit

Sub FizzBuzz()

    Dim bar As UserForm1
    Set bar = New UserForm1
    bar.FooBar = ""FooBar""

    Dim foo As UserForm
    Set foo = New UserForm1
    foo.FooBar = ""BarFoo""

End Sub
";

            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("testproject", ProjectProtection.Unprotected);
            projectBuilder.MockUserFormBuilder("UserForm1", userForm1Code).AddFormToProjectBuilder()
                .AddComponent("ReferencingModule", ComponentType.StandardModule, analyzedCode)
                .AddReference(ReferenceLibrary.MsForms);

            vbeBuilder.AddProject(projectBuilder.Build());
            var vbe = vbeBuilder.Build();

            Assert.IsTrue(InspectionResults(vbe.Object).Any());
        }

        [Test]
        [Ignore("Test concurrency issue. Only passes if run individually.")]
        [Category("Inspections")]
        public void MemberNotOnInterface_DoesNotReturnResult_ControlObject()
        {
            const string inputCode =
                @"Sub Foo(bar As MSForms.TextBox)
    Debug.Print bar.Left
End Sub";

            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("testproject", ProjectProtection.Unprotected);
            projectBuilder.MockUserFormBuilder("UserForm1", inputCode).AddFormToProjectBuilder()
                .AddReference(ReferenceLibrary.MsForms);

            vbeBuilder.AddProject(projectBuilder.Build());
            var vbe = vbeBuilder.Build();

            Assert.AreEqual(0, InspectionResults(vbe.Object).Count());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new MemberNotOnInterfaceInspection(state);
        }
    }
}
