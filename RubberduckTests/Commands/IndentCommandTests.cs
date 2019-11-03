using NUnit.Framework;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Commands
{
    [TestFixture]
    public class IndentCommandTests
    {
        [Category("Commands")]
        [Test]
        public void AddNoIndentAnnotation()
        {
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule("", out component, Selection.Home);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using (state)
            {
                var noIndentAnnotationCommand = MockIndenter.ArrangeNoIndentAnnotationCommand(vbe, state, rewritingManager);
                noIndentAnnotationCommand.Execute(null);

                Assert.AreEqual("'@NoIndent\r\n", component.CodeModule.Content());
            }
        }

        [Category("Commands")]
        [Test]
        public void AddNoIndentAnnotation_ModuleContainsCode()
        {
            var input =
                @"Option Explicit
Public Foo As Boolean

Sub Foo()
End Sub";

            var expected =
                @"'@NoIndent
Option Explicit
Public Foo As Boolean

Sub Foo()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(input, out component, Selection.Home);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using (state)
            {
                var noIndentAnnotationCommand = MockIndenter.ArrangeNoIndentAnnotationCommand(vbe, state, rewritingManager);
                noIndentAnnotationCommand.Execute(null);

                Assert.AreEqual(expected, component.CodeModule.Content());
            }
        }

        [Category("Commands")]
        [Test]
        public void AddNoIndentAnnotation_CanExecute_NullActiveCodePane()
        {
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule("", out component);
            vbe.Setup(v => v.ActiveCodePane).Returns((ICodePane)null);

            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using (state)
            {
                var noIndentAnnotationCommand = MockIndenter.ArrangeNoIndentAnnotationCommand(vbe, state, rewritingManager);
                Assert.IsFalse(noIndentAnnotationCommand.CanExecute(null));
            }
        }

        [Category("Commands")]
        [Test]
        public void AddNoIndentAnnotation_CanExecute_ModuleAlreadyHasAnnotation()
        {
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule("'@NoIndent\r\n", out component);
            vbe.Setup(v => v.ActiveCodePane).Returns((ICodePane)null);

            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using (state)
            {
                var noIndentAnnotationCommand = MockIndenter.ArrangeNoIndentAnnotationCommand(vbe, state, rewritingManager);
                Assert.IsFalse(noIndentAnnotationCommand.CanExecute(null));
            }
        }

        [Category("Commands")]
        [Test]
        public void IndentModule_IndentsModule()
        {
            var input =
                @"    Option Explicit   ' at least I used it...
    Sub InverseIndent()
Dim d As Boolean
Dim s As Integer

    End Sub

   Sub RandomIndent()
Dim d As Boolean
            Dim s As Integer

     End Sub
";

            var expected =
                @"Option Explicit                                  ' at least I used it...

Sub InverseIndent()
    Dim d As Boolean
    Dim s As Integer

End Sub

Sub RandomIndent()
    Dim d As Boolean
    Dim s As Integer

End Sub
";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("Proj1", ProjectProtection.Unprotected)
                .AddComponent("Comp1", ComponentType.ClassModule, input)
                .AddComponent("Comp2", ComponentType.ClassModule, input)
                .Build();

            var vbe = builder.AddProject(project).Build();
            vbe.Setup(s => s.ActiveCodePane).Returns(project.Object.VBComponents["Comp2"].CodeModule.CodePane);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var indentCommand = MockIndenter.ArrangeIndentCurrentModuleCommand(vbe, state);
                indentCommand.Execute(null);

                var module1 = project.Object.VBComponents["Comp1"].CodeModule;
                var module2 = project.Object.VBComponents["Comp2"].CodeModule;

                Assert.AreEqual(input, module1.Content());
                Assert.AreEqual(expected, module2.Content());
            }
        }

        [Category("Commands")]
        [Test]
        public void IndentModule_CanExecute_NullActiveCodePane()
        {
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule("", out component);
            vbe.Setup(v => v.ActiveCodePane).Returns((ICodePane)null);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var indentCommand = MockIndenter.ArrangeIndentCurrentModuleCommand(vbe, state);
                Assert.IsFalse(indentCommand.CanExecute(null));
            }
        }

        [Category("Commands")]
        [Test]
        public void IndentModule_CanExecute()
        {
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule("", out component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var indentCommand = MockIndenter.ArrangeIndentCurrentModuleCommand(vbe, state);
                Assert.IsTrue(indentCommand.CanExecute(null));
            }
        }
    }
}