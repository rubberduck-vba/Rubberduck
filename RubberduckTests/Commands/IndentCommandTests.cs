using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Commands
{
    [TestClass]
    public class IndentCommandTests
    {
        [TestCategory("Commands")]
        [TestMethod]
        public void AddNoIndentAnnotation()
        {
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule("", out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var noIndentAnnotationCommand = new NoIndentAnnotationCommand(vbe.Object, state);
            noIndentAnnotationCommand.Execute(null);

            Assert.AreEqual("'@NoIndent\r\n", component.CodeModule.Content());
        }

        [TestCategory("Commands")]
        [TestMethod]
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
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(input, out component);
            var state = MockParser.CreateAndParse(vbe.Object);
            
            var noIndentAnnotationCommand = new NoIndentAnnotationCommand(vbe.Object, state);
            noIndentAnnotationCommand.Execute(null);

            Assert.AreEqual(expected, component.CodeModule.Content());
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void AddNoIndentAnnotation_CanExecute_NullActiveCodePane()
        {
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule("", out component);
            vbe.Setup(v => v.ActiveCodePane).Returns((ICodePane)null);

            var state = MockParser.CreateAndParse(vbe.Object);

            var noIndentAnnotationCommand = new NoIndentAnnotationCommand(vbe.Object, state);
            Assert.IsFalse(noIndentAnnotationCommand.CanExecute(null));
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void AddNoIndentAnnotation_CanExecute_ModuleAlreadyHasAnnotation()
        {
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule("'@NoIndent\r\n", out component);
            vbe.Setup(v => v.ActiveCodePane).Returns((ICodePane)null);

            var state = MockParser.CreateAndParse(vbe.Object);

            var noIndentAnnotationCommand = new NoIndentAnnotationCommand(vbe.Object, state);
            Assert.IsFalse(noIndentAnnotationCommand.CanExecute(null));
        }

        [TestCategory("Commands")]
        [TestMethod]
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

            var state = MockParser.CreateAndParse(vbe.Object);

            var indentCommand = new IndentCurrentModuleCommand(vbe.Object, CreateIndenter(vbe.Object), state);
            indentCommand.Execute(null);

            var module1 = project.Object.VBComponents["Comp1"].CodeModule;
            var module2 = project.Object.VBComponents["Comp2"].CodeModule;

            Assert.AreEqual(input, module1.Content());
            Assert.AreEqual(expected, module2.Content());
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void IndentModule_CanExecute_NullActiveCodePane()
        {
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule("", out component);
            vbe.Setup(v => v.ActiveCodePane).Returns((ICodePane)null);

            var state = MockParser.CreateAndParse(vbe.Object);

            var indentCommand = new IndentCurrentModuleCommand(vbe.Object, CreateIndenter(vbe.Object), state);
            Assert.IsFalse(indentCommand.CanExecute(null));
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void IndentModule_CanExecute()
        {
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule("", out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var indentCommand = new IndentCurrentModuleCommand(vbe.Object, CreateIndenter(vbe.Object), state);
            Assert.IsTrue(indentCommand.CanExecute(null));
        }

        private static IIndenter CreateIndenter(IVBE vbe)
        {
            return new Indenter(vbe, () => Settings.IndenterSettingsTests.GetMockIndenterSettings());
        }
    }
}