using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class UseSetKeywordForObjectAssignmentQuickFixTests : QuickFixTestBase
    {
        [Test]
        [Category("QuickFixes")]
        public void ObjectVariableNotSet_ReplacesExplicitLetKeyword()
        {
            var inputCode = @"
Private Sub TextBox1_Change()
    Dim foo As Range
    Set foo = Range(""A1"")
    Let foo.Font = Range(""B1"").Font
End Sub
";
            var expectedCode = @"
Private Sub TextBox1_Change()
    Dim foo As Range
    Set foo = Range(""A1"")
    Set foo.Font = Range(""B1"").Font
End Sub
";
            var actualCode = ApplyQuickFixToAllInspectionResults(inputCode, state => new ObjectVariableNotSetInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ObjectVariableNotSet_PlacesKeywordBeforeMemberCall()
        {
            var inputCode = @"
Private Sub TextBox1_Change()
    Dim foo As Range
    Set foo = Range(""A1"")
    foo.Font = Range(""B1"").Font
End Sub
";
            var expectedCode = @"
Private Sub TextBox1_Change()
    Dim foo As Range
    Set foo = Range(""A1"")
    Set foo.Font = Range(""B1"").Font
End Sub
";
            var actualCode = ApplyQuickFixToAllInspectionResults(inputCode, state => new ObjectVariableNotSetInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ObjectVariableNotSet_ForFunctionAssignment_ReturnsResult()
        {
            var inputCode =
                @"
Private Function ReturnObject(ByVal source As Object) As Class1
    ReturnObject = source
End Function";
            var expectedCode =
                @"
Private Function ReturnObject(ByVal source As Object) As Class1
    Set ReturnObject = source
End Function";

            var actualCode = ApplyQuickFixToAllInspectionResults(inputCode, state => new ObjectVariableNotSetInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ObjectVariableNotSet_ForPropertyGetAssignment_ReturnsResults()
        {
            var inputCode = @"
Private m_example As MyObject
Public Property Get Example() As MyObject
    Example = m_example
End Property
";
            var expectedCode =
                @"
Private m_example As MyObject
Public Property Get Example() As MyObject
    Set Example = m_example
End Property
";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ObjectVariableNotSetInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("Inspections")]
        public void BothSidesOfAssignmentHaveDefaultMemberAccess_NoExplicitLet_QuickFixWorks()
        {
            var class1Code = @"
Public Property Get Baz() As Long
Attribute Baz.VB_UserMemId = 0
End Property

Public Property Let Baz(RHS As Long)
End Property
";

            var moduleCode = $@"
Private Sub Bar() 
    Dim cls1 As Class1
    Dim cls2 As Class1
    Set cls1 = New Class1
    Set cls2 = New Class1
    cls2 = cls1
End Sub
";

            var expectedModuleCode = $@"
Private Sub Bar() 
    Dim cls1 As Class1
    Dim cls2 As Class1
    Set cls1 = New Class1
    Set cls2 = New Class1
    Set cls2 = cls1
End Sub
";

            var vbe = MockVbeBuilder.BuildFromModules(
                ("Class1", class1Code, ComponentType.ClassModule),
                ("Module1", moduleCode, ComponentType.StandardModule));

            var actualModuleCode = ApplyQuickFixToAllInspectionResults(vbe.Object, "Module1", state => new SuspiciousLetAssignmentInspection(state));
            Assert.AreEqual(expectedModuleCode, actualModuleCode);
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new UseSetKeywordForObjectAssignmentQuickFix();
        }

        protected override IVBE TestVbe(string code, out IVBComponent component)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, code)
                .AddReference(ReferenceLibrary.Excel)
                .Build();

            var vbe = builder.AddProject(project).Build().Object;
            component = project.Object.VBComponents[0];
            return vbe;
        }

    }
}
