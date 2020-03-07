using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class PassParameterByValueQuickFixTests : QuickFixTestBase
    {
        [Test]
        [Category("QuickFixes")]
        public void ParameterCanBeByVal_QuickFixWorks_SubNameStartsWithParamName()
        {
            const string inputCode =
                @"Sub foo(f)
End Sub";

            const string expectedCode =
                @"Sub foo(ByVal f)
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ParameterCanBeByValInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ParameterCanBeByVal_QuickFixWorks_PassedByUnspecified()
        {
            const string inputCode =
                @"Sub Foo(arg1 As String)
End Sub";

            const string expectedCode =
                @"Sub Foo(ByVal arg1 As String)
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ParameterCanBeByValInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ParameterCanBeByVal_QuickFixWorks_PassedByRef()
        {
            const string inputCode =
                @"Sub Foo(ByRef arg1 As String)
End Sub";

            const string expectedCode =
                @"Sub Foo(ByVal arg1 As String)
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ParameterCanBeByValInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ParameterCanBeByVal_QuickFixWorks_PassedByUnspecified_MultilineParameter()
        {
            const string inputCode =
                @"Sub Foo( _
arg1 As String)
End Sub";

            const string expectedCode =
                @"Sub Foo( _
ByVal arg1 As String)
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ParameterCanBeByValInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ParameterCanBeByVal_QuickFixWorks_PassedByRef_MultilineParameter()
        {
            const string inputCode =
                @"Sub Foo(ByRef _
arg1 As String)
End Sub";

            const string expectedCode =
                @"Sub Foo(ByVal _
arg1 As String)
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ParameterCanBeByValInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ParameterCanBeByVal_InterfaceMember_MultipleParams_OneCanBeByVal_QuickFixWorks()
        {
            //Input
            const string inputCode1 =
                @"Public Sub DoSomething(ByRef a As Integer, ByRef b As Integer)
End Sub";
            const string inputCode2 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByRef a As Integer, ByRef b As Integer)
    b = 42
End Sub";
            const string inputCode3 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByRef a As Integer, ByRef b As Integer)
End Sub";

            //Expected
            const string expectedCode1 =
                @"Public Sub DoSomething(ByVal a As Integer, ByRef b As Integer)
End Sub";
            const string expectedCode2 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer, ByRef b As Integer)
    b = 42
End Sub";
            const string expectedCode3 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer, ByRef b As Integer)
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .AddComponent("Class2", ComponentType.ClassModule, inputCode3)
                .Build();

            var component1 = project.Object.VBComponents["IClass1"];
            var component2 = project.Object.VBComponents["Class1"];
            var component3 = project.Object.VBComponents["Class2"];
            var vbe = builder.AddProject(project).Build();

            var (state, rewriteManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using (state)
            {

                var inspection = new ParameterCanBeByValInspection(state);
                var rewriteSession = rewriteManager.CheckOutCodePaneSession();

                new PassParameterByValueQuickFix(state).Fix(inspection.GetInspectionResults(CancellationToken.None).First(), rewriteSession);

                var actualCode1 = rewriteSession.CheckOutModuleRewriter(component1.QualifiedModuleName).GetText();
                var actualCode2 = rewriteSession.CheckOutModuleRewriter(component2.QualifiedModuleName).GetText();
                var actualCode3 = rewriteSession.CheckOutModuleRewriter(component3.QualifiedModuleName).GetText();

                Assert.AreEqual(expectedCode1, actualCode1);
                Assert.AreEqual(expectedCode2, actualCode2);
                Assert.AreEqual(expectedCode3, actualCode3);
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ParameterCanBeByVal_EventMember_MultipleParams_OneCanBeByVal_QuickFixWorks()
        {
            //Input
            const string inputCode1 =
                @"Public Event Foo(ByRef a As Integer, ByRef b As Integer)";
            const string inputCode2 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByRef a As Integer, ByRef b As Integer)
    a = 42
End Sub";
            const string inputCode3 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByRef a As Integer, ByRef b As Integer)
End Sub";

            //Expected
            const string expectedCode1 =
                @"Public Event Foo(ByRef a As Integer, ByVal b As Integer)";
            const string expectedCode2 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByRef a As Integer, ByVal b As Integer)
    a = 42
End Sub";
            const string expectedCode3 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByRef a As Integer, ByVal b As Integer)
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class2", ComponentType.ClassModule, inputCode2)
                .AddComponent("Class3", ComponentType.ClassModule, inputCode3)
                .Build();

            var component1 = project.Object.VBComponents["Class1"];
            var component2 = project.Object.VBComponents["Class2"];
            var component3 = project.Object.VBComponents["Class3"];
            var vbe = builder.AddProject(project).Build();

            var (state, rewriteManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using (state)
            {

                var inspection = new ParameterCanBeByValInspection(state);
                var rewriteSession = rewriteManager.CheckOutCodePaneSession();

                new PassParameterByValueQuickFix(state).Fix(inspection.GetInspectionResults(CancellationToken.None).First(), rewriteSession);

                var actualCode1 = rewriteSession.CheckOutModuleRewriter(component1.QualifiedModuleName).GetText();
                var actualCode2 = rewriteSession.CheckOutModuleRewriter(component2.QualifiedModuleName).GetText();
                var actualCode3 = rewriteSession.CheckOutModuleRewriter(component3.QualifiedModuleName).GetText();

                Assert.AreEqual(expectedCode1, actualCode1);
                Assert.AreEqual(expectedCode2, actualCode2);
                Assert.AreEqual(expectedCode3, actualCode3);
            }
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/2408
        [Test]
        [Category("QuickFixes")]
        public void ParameterCanBeByVal_QuickFixWithOptionalWorks()
        {
            const string inputCode =
                @"Sub Test(Optional foo As String = ""bar"")
    Debug.Print foo
End Sub";

            const string expectedCode =
                @"Sub Test(Optional ByVal foo As String = ""bar"")
    Debug.Print foo
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ParameterCanBeByValInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/2408
        [Test]
        [Category("QuickFixes")]
        public void ParameterCanBeByVal_QuickFixWithOptionalByRefWorks()
        {
            const string inputCode =
                @"Sub Test(Optional ByRef foo As String = ""bar"")
    Debug.Print foo
End Sub";

            const string expectedCode =
                @"Sub Test(Optional ByVal foo As String = ""bar"")
    Debug.Print foo
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ParameterCanBeByValInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/2408
        [Test]
        [Category("QuickFixes")]
        public void ParameterCanBeByVal_QuickFixWithOptional_LineContinuationsWorks()
        {
            const string inputCode =
                @"Sub foo(Optional _
  ByRef _
  foo _
  As _
  Byte _
  )
  Debug.Print foo
End Sub";

            const string expectedCode =
                @"Sub foo(Optional _
  ByVal _
  foo _
  As _
  Byte _
  )
  Debug.Print foo
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ParameterCanBeByValInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new PassParameterByValueQuickFix(state);
        }
    }
}
