using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.VBA;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.Application;
using Rubberduck.VBEditor.Events;
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
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule("", out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));
            var module = component.CodeModule;
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }
            
            var noIndentAnnotationCommand = new NoIndentAnnotationCommand(vbe.Object, parser.State);
            noIndentAnnotationCommand.Execute(null);

            Assert.AreEqual("'@NoIndent\r\n", module.Content());
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

            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(input, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var module = component.CodeModule;
            var noIndentAnnotationCommand = new NoIndentAnnotationCommand(vbe.Object, parser.State);
            noIndentAnnotationCommand.Execute(null);

            Assert.AreEqual(expected, module.Content());
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void AddNoIndentAnnotation_CanExecute_NullActiveCodePane()
        {
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule("", out component);
            vbe.Setup(v => v.ActiveCodePane).Returns((ICodePane) null);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var noIndentAnnotationCommand = new NoIndentAnnotationCommand(vbe.Object, parser.State);
            Assert.IsFalse(noIndentAnnotationCommand.CanExecute(null));
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void AddNoIndentAnnotation_CanExecute_ModuleAlreadyHasAnnotation()
        {
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule("'@NoIndent\r\n", out component);
            vbe.Setup(v => v.ActiveCodePane).Returns((ICodePane)null);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var noIndentAnnotationCommand = new NoIndentAnnotationCommand(vbe.Object, parser.State);
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

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var indentCommand = new IndentCurrentModuleCommand(vbe.Object, CreateIndenter(vbe.Object), parser.State);
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
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule("", out component);
            vbe.Setup(v => v.ActiveCodePane).Returns((ICodePane) null);

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var indentCommand = new IndentCurrentModuleCommand(vbe.Object, CreateIndenter(vbe.Object), parser.State);
            Assert.IsFalse(indentCommand.CanExecute(null));
        }

        [TestCategory("Commands")]
        [TestMethod]
        public void IndentModule_CanExecute()
        {
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule("", out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var indentCommand = new IndentCurrentModuleCommand(vbe.Object, CreateIndenter(vbe.Object), parser.State);
            Assert.IsTrue(indentCommand.CanExecute(null));
        }

        private static IIndenter CreateIndenter(IVBE vbe)
        {
            return new Indenter(vbe, () => Settings.IndenterSettingsTests.GetMockIndenterSettings());
        }
    }
}