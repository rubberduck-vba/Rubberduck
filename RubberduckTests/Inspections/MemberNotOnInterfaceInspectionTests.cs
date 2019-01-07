using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class MemberNotOnInterfaceInspectionTests
    {
        private static RubberduckParserState ArrangeParserAndParse(string inputCode, string library = "Scripting")
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("Codez", ComponentType.StandardModule, inputCode)
                .AddReference(library,
                    library.Equals("Scripting") ? MockVbeBuilder.LibraryPathScripting : MockVbeBuilder.LibraryPathMsExcel,
                    1,
                    library.Equals("Scripting") ? 0 : 8,
                    true)
                .Build();

            var vbe = builder.AddProject(project).Build();

            return MockParser.CreateAndParse(vbe.Object);
        }

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

            using (var state = ArrangeParserAndParse(inputCode))
            {
                var inspection = new MemberNotOnInterfaceInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(1, inspectionResults.Count());
            }
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

            using (var state = ArrangeParserAndParse(inputCode))
            {
                var inspection = new MemberNotOnInterfaceInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void MemberNotOnInterface_ReturnsResult_ApplicationObject()
        {
            const string inputCode =
                @"Sub Foo()
    Application.NonMember
End Sub";

            using (var state = ArrangeParserAndParse(inputCode, "Excel"))
            {
                var inspection = new MemberNotOnInterfaceInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void MemberNotOnInterface_ReturnsResult_UnDeclaredMemberOnParameter()
        {
            const string inputCode =
                @"Sub Foo(dict As Dictionary)
    dict.NonMember
End Sub";

            using (var state = ArrangeParserAndParse(inputCode))
            {
                var inspection = new MemberNotOnInterfaceInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(1, inspectionResults.Count());
            }
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

            using (var state = ArrangeParserAndParse(inputCode))
            {
                var inspection = new MemberNotOnInterfaceInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.IsFalse(inspectionResults.Any());
            }
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

            using (var state = ArrangeParserAndParse(inputCode))
            {
                var inspection = new MemberNotOnInterfaceInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.IsFalse(inspectionResults.Any());
            }
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

            using (var state = ArrangeParserAndParse(inputCode))
            {
                var inspection = new MemberNotOnInterfaceInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(1, inspectionResults.Count());
            }
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

            using (var state = ArrangeParserAndParse(inputCode))
            {
                var inspection = new MemberNotOnInterfaceInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.IsFalse(inspectionResults.Any());
            }
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

            using (var state = ArrangeParserAndParse(inputCode))
            {
                var inspection = new MemberNotOnInterfaceInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void MemberNotOnInterface_DoesNotReturnResult_ProjectReference()
        {
            const string inputCode =
                @"Sub Foo()
    Dim dict As Scripting.Dictionary
End Sub";

            using (var state = ArrangeParserAndParse(inputCode))
            {
                var inspection = new MemberNotOnInterfaceInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.IsFalse(inspectionResults.Any());
            }
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

            using (var state = ArrangeParserAndParse(inputCode))
            {
                var inspection = new MemberNotOnInterfaceInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.IsFalse(inspectionResults.Any());
            }
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

            using (var state = ArrangeParserAndParse(inputCode))
            {
                var inspection = new MemberNotOnInterfaceInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        //See https://github.com/rubberduck-vba/Rubberduck/issues/4308 
        [Test]
        [Category("Inspections")]
        [Ignore("To be unignored in a PR fixing issue 4308.")]
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

            using (var state = ArrangeParserAndParse(inputCode))
            {
                var inspection = new MemberNotOnInterfaceInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.AreEqual(1, inspectionResults.Count());
            }
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

            using (var state = ArrangeParserAndParse(inputCode))
            {
                var inspection = new MemberNotOnInterfaceInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.IsFalse(inspectionResults.Any());
            }
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

            using (var state = ArrangeParserAndParse(inputCode))
            {
                var inspection = new MemberNotOnInterfaceInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.IsFalse(inspectionResults.Any());
            }
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
                .AddReference("MSForms", MockVbeBuilder.LibraryPathMsForms, 2, 0, true);

            vbeBuilder.AddProject(projectBuilder.Build());
            var vbe = vbeBuilder.Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new MemberNotOnInterfaceInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.IsTrue(inspectionResults.Any());
            }

        }

        [Test]
        [Ignore("Test concurrency issue. Only passes if run individually.")]
        [Category("Inspections")]
        public void MemberNotOnInterface_DoesNotReturnResult_ControlObject()
        {
            const string inputCode =
                @"Sub Foo(bar as MSForms.TextBox)
    Debug.Print bar.Left
End Sub";

            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder("testproject", ProjectProtection.Unprotected);
            projectBuilder.MockUserFormBuilder("UserForm1", inputCode).AddFormToProjectBuilder()
                .AddReference("MSForms", MockVbeBuilder.LibraryPathMsForms, 2, 0, true);

            vbeBuilder.AddProject(projectBuilder.Build());
            var vbe = vbeBuilder.Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new MemberNotOnInterfaceInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                Assert.IsTrue(!inspectionResults.Any());
            }
        }
    }
}
