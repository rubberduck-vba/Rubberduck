using System.Threading;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.VBA;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEHost;
using RubberduckTests.Mocks;

namespace RubberduckTests.Commands
{
    [TestClass]
    public class IndentCommandTests
    {
        [TestMethod]
        public void AddNoIndentAnnotation()
        {
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule("", out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }
            
            var noIndentAnnotationCommand = new NoIndentAnnotationCommand(vbe.Object, parser.State);
            noIndentAnnotationCommand.Execute(null);

            Assert.AreEqual("'@NoIndent\r\n", component.CodeModule.Lines());
        }

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
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(input, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var noIndentAnnotationCommand = new NoIndentAnnotationCommand(vbe.Object, parser.State);
            noIndentAnnotationCommand.Execute(null);

            Assert.AreEqual(expected, component.CodeModule.Lines());
        }

        [TestMethod]
        public void AddNoIndentAnnotation_CanExecute_NullActiveCodePane()
        {
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule("", out component);
            vbe.Setup(v => v.ActiveCodePane).Returns((CodePane) null);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var noIndentAnnotationCommand = new NoIndentAnnotationCommand(vbe.Object, parser.State);
            Assert.IsFalse(noIndentAnnotationCommand.CanExecute(null));
        }

        [TestMethod]
        public void AddNoIndentAnnotation_CanExecute_ModuleAlreadyHasAnnotation()
        {
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule("'@NoIndent\r\n", out component);
            vbe.Setup(v => v.ActiveCodePane).Returns((CodePane)null);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var noIndentAnnotationCommand = new NoIndentAnnotationCommand(vbe.Object, parser.State);
            Assert.IsFalse(noIndentAnnotationCommand.CanExecute(null));
        }

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
            var project = builder.ProjectBuilder("Proj1", vbext_ProjectProtection.vbext_pp_none)
                .AddComponent("Comp1", vbext_ComponentType.vbext_ct_ClassModule, input)
                .AddComponent("Comp2", vbext_ComponentType.vbext_ct_ClassModule, input)
                .Build();

            var vbe = builder.AddProject(project).Build();
            vbe.Setup(s => s.ActiveCodePane).Returns(project.Object.VBComponents.Item("Comp2").CodeModule.CodePane);

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var indentCommand = new IndentCurrentModuleCommand(vbe.Object, CreateIndenter(vbe.Object));
            indentCommand.Execute(null);

            Assert.AreEqual(input, project.Object.VBComponents.Item("Comp1").CodeModule.Lines());
            Assert.AreEqual(expected, project.Object.VBComponents.Item("Comp2").CodeModule.Lines());
        }

        [TestMethod]
        public void IndentModule_CanExecute_NullActiveCodePane()
        {
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule("", out component);
            vbe.Setup(v => v.ActiveCodePane).Returns((CodePane) null);

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var indentCommand = new IndentCurrentModuleCommand(vbe.Object, CreateIndenter(vbe.Object));
            Assert.IsFalse(indentCommand.CanExecute(null));
        }

        [TestMethod]
        public void IndentModule_CanExecute()
        {
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule("", out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error)
            {
                Assert.Inconclusive("Parser Error");
            }

            var indentCommand = new IndentCurrentModuleCommand(vbe.Object, CreateIndenter(vbe.Object));
            Assert.IsTrue(indentCommand.CanExecute(null));
        }

        private static IIndenter CreateIndenter(VBE vbe)
        {
            var settings = new Mock<IndenterSettings>();
            settings.Setup(s => s.IndentEntireProcedureBody).Returns(true);
            settings.Setup(s => s.IndentFirstCommentBlock).Returns(true);
            settings.Setup(s => s.IndentFirstDeclarationBlock).Returns(true);
            settings.Setup(s => s.AlignCommentsWithCode).Returns(true);
            settings.Setup(s => s.AlignContinuations).Returns(true);
            settings.Setup(s => s.IgnoreOperatorsInContinuations).Returns(true);
            settings.Setup(s => s.IndentCase).Returns(false);
            settings.Setup(s => s.ForceDebugStatementsInColumn1).Returns(false);
            settings.Setup(s => s.ForceCompilerDirectivesInColumn1).Returns(false);
            settings.Setup(s => s.IndentCompilerDirectives).Returns(true);
            settings.Setup(s => s.AlignDims).Returns(false);
            settings.Setup(s => s.AlignDimColumn).Returns(15);
            settings.Setup(s => s.EnableUndo).Returns(true);
            settings.Setup(s => s.EndOfLineCommentStyle).Returns(EndOfLineCommentStyle.AlignInColumn);
            settings.Setup(s => s.EndOfLineCommentColumnSpaceAlignment).Returns(50);
            settings.Setup(s => s.IndentSpaces).Returns(4);

            return new Indenter(vbe, () => new IndenterSettings());
        }
    }
}