using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Antlr4.Runtime;
using NUnit.Framework;
using Moq;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SourceCodeHandling;
using RubberduckTests.Mocks;

// ReSharper disable InvokeAsExtensionMethod
namespace RubberduckTests.PostProcessing
{
    [TestFixture]
    public class ModuleRewriterTests
    {
        [Test]
        [Category("TokenStreamRewriter")]
        public void RewriteClearsEntireModule()
        {
            var codeModule = new Mock<ICodeModule>();
            codeModule.Setup(m => m.Clear());

            var module = new QualifiedModuleName("TestProject", string.Empty,"TestModule");
            var projectsProvider = TestProvider(module, codeModule.Object);
            var codePaneSourceHandler = new CodePaneSourceCodeHandler(projectsProvider);
            var tokenStream = new CommonTokenStream(new ListTokenSource(new List<IToken>()));

            var sut = new ModuleRewriter(module, tokenStream, codePaneSourceHandler);
            sut.InsertAfter(0, "test");

            if (!sut.IsDirty)
            {
                sut.InsertBefore(0, "foo");
            }
            sut.Rewrite();

            codeModule.Verify(m => m.Clear());
        }

        [Test]
        [Category("TokenStreamRewriter")]
        public void RewriteDoesNotRewriteIfNotDirty()
        {
            var codeModule = new Mock<ICodeModule>();
            codeModule.Setup(m => m.Content()).Returns(string.Empty);
            codeModule.Setup(m => m.Clear());

            var module = new QualifiedModuleName("TestProject", string.Empty, "TestModule");
            var projectsProvider = TestProvider(module, codeModule.Object);
            var codePaneSourceHandler = new CodePaneSourceCodeHandler(projectsProvider);
            var tokenStream = new CommonTokenStream(new ListTokenSource(new List<IToken>()));

            var sut = new ModuleRewriter(module, tokenStream, codePaneSourceHandler);

            sut.Rewrite();
            codeModule.Verify(m => m.Clear(), Times.Never);
        }

        [Test]
        [Category("TokenStreamRewriter")]
        public void RewriterInsertsRewriterOutputAtLine1()
        {
            const string content = @"Option Explicit";
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(content, out component).Object;

            using (var state = MockParser.CreateAndParse(vbe))
            {
                if (state.Status != ParserState.Ready)
                {
                    Assert.Inconclusive("Parser isn't ready. Test cannot proceed.");
                }

                var rewriter = state.GetRewriter(component);
                rewriter.Rewrite();

                Assert.AreEqual(content, rewriter.GetText());
            }
        }

        [Test]
        [Category("TokenStreamRewriter")]
        public void RemovesModuleVariableDeclarationStatement()
        {
            const string expected = @"
";
            const string content = @"
Private foo As String
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(content, out _).Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                if (state.Status != ParserState.Ready)
                {
                    Assert.Inconclusive("Parser isn't ready. Test cannot proceed.");
                }

                var declarations = state.AllUserDeclarations;
                var target = declarations.SingleOrDefault(d => d.DeclarationType == DeclarationType.Variable);
                if (target == null)
                {
                    Assert.Inconclusive("No variable was found in test code.");
                }

                var rewriter = state.GetRewriter(target);
                rewriter.Remove(target);

                Assert.AreEqual(expected, rewriter.GetText());
            }
        }

        [Test]
        [Category("TokenStreamRewriter")]
        public void RemovesModuleConstantDeclarationStatement()
        {
            const string expected = @"
";
            const string content = @"
Private Const foo As String = ""Something""
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(content, out _).Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                if (state.Status != ParserState.Ready)
                {
                    Assert.Inconclusive("Parser isn't ready. Test cannot proceed.");
                }

                var declarations = state.AllUserDeclarations;
                var target = declarations.SingleOrDefault(d => d.DeclarationType == DeclarationType.Constant);
                if (target == null)
                {
                    Assert.Inconclusive("No constant was found in test code.");
                }

                var rewriter = state.GetRewriter(target);
                rewriter.Remove(target);

                Assert.AreEqual(expected, rewriter.GetText());
            }
        }

        [Test]
        [Category("TokenStreamRewriter")]
        public void RemovesLocalVariableDeclarationStatement()
        {
            const string expected = @"
Sub DoSomething()
End Sub
";
            const string content = @"
Sub DoSomething()
Dim foo As String
End Sub
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(content, out _).Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                if (state.Status != ParserState.Ready)
                {
                    Assert.Inconclusive("Parser isn't ready. Test cannot proceed.");
                }

                var declarations = state.AllUserDeclarations;
                var target = declarations.SingleOrDefault(d => d.DeclarationType == DeclarationType.Variable);
                if (target == null)
                {
                    Assert.Inconclusive("No variable was found in test code.");
                }

                var rewriter = state.GetRewriter(target);
                rewriter.Remove(target);

                Assert.AreEqual(expected, rewriter.GetText());
            }
        }

        [Test]
        [Category("TokenStreamRewriter")]
        public void RemovesLocalConstantDeclarationStatement()
        {
            const string expected = @"
Sub DoSomething()
End Sub
";
            const string content = @"
Sub DoSomething()
Const foo As String = ""Something""
End Sub
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(content, out _).Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                if (state.Status != ParserState.Ready)
                {
                    Assert.Inconclusive("Parser isn't ready. Test cannot proceed.");
                }

                var declarations = state.AllUserDeclarations;
                var target = declarations.SingleOrDefault(d => d.DeclarationType == DeclarationType.Constant);
                if (target == null)
                {
                    Assert.Inconclusive("No constant was found in test code.");
                }

                var rewriter = state.GetRewriter(target);
                rewriter.Remove(target);

                var rewrittenCode = rewriter.GetText();
                Assert.AreEqual(expected, rewrittenCode);
            }
        }

        [Test]
        [Category("TokenStreamRewriter")]
        public void RemovesSingleParameterDeclaration()
        {
            const string expected = @"
Sub DoSomething()
End Sub
";
            const string content = @"
Sub DoSomething(ByVal foo As Long)
End Sub
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(content, out _).Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                if (state.Status != ParserState.Ready)
                {
                    Assert.Inconclusive("Parser isn't ready. Test cannot proceed.");
                }

                var declarations = state.AllUserDeclarations;
                var target = declarations.SingleOrDefault(d => d.DeclarationType == DeclarationType.Parameter);
                if (target == null)
                {
                    Assert.Inconclusive("No parameter was found in test code.");
                }

                var rewriter = state.GetRewriter(target);
                rewriter.Remove(target);

                Assert.AreEqual(expected, rewriter.GetText());
            }
        }

        [Test]
        [Category("TokenStreamRewriter")]
        public void RemovesEventParameterDeclaration()
        {
            const string expected = @"
Public Event SomeEvent()
";
            const string content = @"
Public Event SomeEvent(ByVal foo As Long)
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(content, out _).Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                if (state.Status != ParserState.Ready)
                {
                    Assert.Inconclusive("Parser isn't ready. Test cannot proceed.");
                }

                var declarations = state.AllUserDeclarations;
                var target = declarations.SingleOrDefault(d => d.DeclarationType == DeclarationType.Parameter);
                if (target == null)
                {
                    Assert.Inconclusive("No parameter was found in test code.");
                }

                var rewriter = state.GetRewriter(target);
                rewriter.Remove(target);

                Assert.AreEqual(expected, rewriter.GetText());
            }
        }

        [Test]
        [Category("TokenStreamRewriter")]
        public void RemovesDeclareFunctionParameterDeclaration()
        {
            const string expected = @"
Declare PtrSafe Function Foo Lib ""Z"" Alias ""Y"" () As Long
";
            const string content = @"
Declare PtrSafe Function Foo Lib ""Z"" Alias ""Y"" (ByVal bar As Long) As Long
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(content, out _).Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                if (state.Status != ParserState.Ready)
                {
                    Assert.Inconclusive("Parser isn't ready. Test cannot proceed.");
                }

                var declarations = state.AllUserDeclarations;
                var target = declarations.SingleOrDefault(d => d.DeclarationType == DeclarationType.Parameter);
                if (target == null)
                {
                    Assert.Inconclusive("No parameter was found in test code.");
                }

                var rewriter = state.GetRewriter(target);
                rewriter.Remove(target);

                Assert.AreEqual(expected, rewriter.GetText());
            }
        }

        [Test]
        [Category("TokenStreamRewriter")]
        public void RemovesFirstVariableInStatement()
        {
            const string expected = @"
Sub DoSomething()
Dim bar As Integer
End Sub
";
            const string content = @"
Sub DoSomething()
Dim foo As String, bar As Integer
End Sub
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(content, out _).Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                if (state.Status != ParserState.Ready)
                {
                    Assert.Inconclusive("Parser isn't ready. Test cannot proceed.");
                }

                var declarations = state.AllUserDeclarations;
                var target = declarations.SingleOrDefault(d => d.IdentifierName == "foo" && d.DeclarationType == DeclarationType.Variable);
                if (target == null)
                {
                    Assert.Inconclusive("No 'foo' variable was found in test code.");
                }

                var rewriter = state.GetRewriter(target);
                rewriter.Remove(target);

                Assert.AreEqual(expected, rewriter.GetText());
            }
        }

        [Test]
        [Category("TokenStreamRewriter")]
        public void RemovesFirstConstantInStatement()
        {
            const string expected = @"
Sub DoSomething()
Const bar As Integer = 42
End Sub
";
            const string content = @"
Sub DoSomething()
Const foo As String = ""Something"", bar As Integer = 42
End Sub
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(content, out _).Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                if (state.Status != ParserState.Ready)
                {
                    Assert.Inconclusive("Parser isn't ready. Test cannot proceed.");
                }

                var declarations = state.AllUserDeclarations;
                var target = declarations.SingleOrDefault(d => d.IdentifierName == "foo" && d.DeclarationType == DeclarationType.Constant);
                if (target == null)
                {
                    Assert.Inconclusive("No 'foo' constant was found in test code.");
                }

                var rewriter = state.GetRewriter(target);
                rewriter.Remove(target);

                Assert.AreEqual(expected, rewriter.GetText());
            }
        }

        [Test]
        [Category("TokenStreamRewriter")]
        public void RemovesFirstParameterInSignature()
        {
            const string expected = @"
Sub DoSomething(ByVal bar As Long)
End Sub
";
            const string content = @"
Sub DoSomething(ByVal foo As Long, ByVal bar As Long)
End Sub
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(content, out _).Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                if (state.Status != ParserState.Ready)
                {
                    Assert.Inconclusive("Parser isn't ready. Test cannot proceed.");
                }

                var declarations = state.AllUserDeclarations;
                var target = declarations.SingleOrDefault(d => d.IdentifierName == "foo" && d.DeclarationType == DeclarationType.Parameter);
                if (target == null)
                {
                    Assert.Inconclusive("No 'foo' parameter was found in test code.");
                }

                var rewriter = state.GetRewriter(target);
                rewriter.Remove(target);

                Assert.AreEqual(expected, rewriter.GetText());
            }
        }

        [Test]
        [Category("TokenStreamRewriter")]
        public void RemovesLastVariableInStatement()
        {
            const string expected = @"
Sub DoSomething()
Dim foo As String
End Sub
";
            const string content = @"
Sub DoSomething()
Dim foo As String, bar As Integer
End Sub
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(content, out _).Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                if (state.Status != ParserState.Ready)
                {
                    Assert.Inconclusive("Parser isn't ready. Test cannot proceed.");
                }

                var declarations = state.AllUserDeclarations;
                var target = declarations.SingleOrDefault(d => d.IdentifierName == "bar" && d.DeclarationType == DeclarationType.Variable);
                if (target == null)
                {
                    Assert.Inconclusive("No variable was found in test code.");
                }

                var rewriter = state.GetRewriter(target);
                rewriter.Remove(target);

                Assert.AreEqual(expected, rewriter.GetText());
            }
        }

        [Test]
        [Category("TokenStreamRewriter")]
        public void RemovesLastParameterInSignature()
        {
            const string expected = @"
Sub DoSomething(ByVal foo As Long)
End Sub
";
            const string content = @"
Sub DoSomething(ByVal foo As Long, ByVal bar As Long)
End Sub
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(content, out _).Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                if (state.Status != ParserState.Ready)
                {
                    Assert.Inconclusive("Parser isn't ready. Test cannot proceed.");
                }

                var declarations = state.AllUserDeclarations;
                var target = declarations.SingleOrDefault(d => d.IdentifierName == "bar" && d.DeclarationType == DeclarationType.Parameter);
                if (target == null)
                {
                    Assert.Inconclusive("No parameter was found in test code.");
                }

                var rewriter = state.GetRewriter(target);
                rewriter.Remove(target);

                Assert.AreEqual(expected, rewriter.GetText());
            }
        }

        [Test]
        [Category("TokenStreamRewriter")]
        public void RemovesLastConstantInStatement()
        {
            const string expected = @"
Sub DoSomething()
Const foo As String = ""Something""
End Sub
";
            const string content = @"
Sub DoSomething()
Const foo As String = ""Something"", bar As Integer = 42
End Sub
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(content, out _).Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                if (state.Status != ParserState.Ready)
                {
                    Assert.Inconclusive("Parser isn't ready. Test cannot proceed.");
                }

                var declarations = state.AllUserDeclarations;
                var target = declarations.SingleOrDefault(d => d.IdentifierName == "bar" && d.DeclarationType == DeclarationType.Constant);
                if (target == null)
                {
                    Assert.Inconclusive("No 'bar' constant was found in test code.");
                }

                var rewriter = state.GetRewriter(target);
                rewriter.Remove(target);

                Assert.AreEqual(expected, rewriter.GetText());
            }
        }

        [Test]
        [Category("TokenStreamRewriter")]
        public void RemovesModuleVariableDeclarationWithLineContinuations()
        {
            const string expected = @"
";
            const string content = @"
Private foo _
  As _
    String
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(content, out _).Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                if (state.Status != ParserState.Ready)
                {
                    Assert.Inconclusive("Parser isn't ready. Test cannot proceed.");
                }

                var declarations = state.AllUserDeclarations;
                var target = declarations.SingleOrDefault(d => d.DeclarationType == DeclarationType.Variable);
                if (target == null)
                {
                    Assert.Inconclusive("No variable was found in test code.");
                }

                var rewriter = state.GetRewriter(target);
                rewriter.Remove(target);

                Assert.AreEqual(expected, rewriter.GetText());
            }
        }

        [Test]
        [Category("TokenStreamRewriter")]
        public void RemovesModuleConstantDeclarationWithLineContinuations()
        {
            const string expected = @"
";
            const string content = @"
Private Const foo _
  As String = _
  ""Something""
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(content, out _).Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                if (state.Status != ParserState.Ready)
                {
                    Assert.Inconclusive("Parser isn't ready. Test cannot proceed.");
                }

                var declarations = state.AllUserDeclarations;
                var target = declarations.SingleOrDefault(d => d.DeclarationType == DeclarationType.Constant);
                if (target == null)
                {
                    Assert.Inconclusive("No constant was found in test code.");
                }

                var rewriter = state.GetRewriter(target);
                rewriter.Remove(target);

                Assert.AreEqual(expected, rewriter.GetText());
            }
        }

        [Test]
        [Category("TokenStreamRewriter")]
        public void RemovesFirstVariableInDeclarationList()
        {
            const string content = @"
Private foo As String, bar As Long
";
            const string expected = @"
Private bar As Long
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(content, out _).Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                if (state.Status != ParserState.Ready)
                {
                    Assert.Inconclusive("Parser isn't ready. Test cannot proceed.");
                }

                var declarations = state.AllUserDeclarations;
                var target = declarations.SingleOrDefault(d => d.IdentifierName == "foo" && d.DeclarationType == DeclarationType.Variable);
                if (target == null)
                {
                    Assert.Inconclusive("Target variable was not found in test code.");
                }

                var rewriter = state.GetRewriter(target);
                rewriter.Remove(target);

                Assert.AreEqual(expected, rewriter.GetText());
            }
        }

        [Test]
        [Category("TokenStreamRewriter")]
        public void RemovesFirstConstantInDeclarationList()
        {
            const string content = @"
Private Const foo As String = ""Something"", bar As Long = 42
";
            const string expected = @"
Private Const bar As Long = 42
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(content, out _).Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                if (state.Status != ParserState.Ready)
                {
                    Assert.Inconclusive("Parser isn't ready. Test cannot proceed.");
                }

                var declarations = state.AllUserDeclarations;
                var target = declarations.SingleOrDefault(d => d.IdentifierName == "foo" && d.DeclarationType == DeclarationType.Constant);
                if (target == null)
                {
                    Assert.Inconclusive("Target constant was not found in test code.");
                }

                var rewriter = state.GetRewriter(target);
                rewriter.Remove(target);

                Assert.AreEqual(expected, rewriter.GetText());
            }
        }

        [Test]
        [Category("TokenStreamRewriter")]
        public void RemovesLastVariableInDeclarationList()
        {
            const string content = @"
Private foo As String, bar As Long
";
            const string expected = @"
Private foo As String
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(content, out _).Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                if (state.Status != ParserState.Ready)
                {
                    Assert.Inconclusive("Parser isn't ready. Test cannot proceed.");
                }

                var declarations = state.AllUserDeclarations;
                var target = declarations.SingleOrDefault(d => d.IdentifierName == "bar" && d.DeclarationType == DeclarationType.Variable);
                if (target == null)
                {
                    Assert.Inconclusive("Target variable was not found in test code.");
                }

                var rewriter = state.GetRewriter(target);
                rewriter.Remove(target);

                Assert.AreEqual(expected, rewriter.GetText());
            }
        }

        [Test]
        [Category("TokenStreamRewriter")]
        public void RemovesLastConstantInDeclarationList()
        {
            const string content = @"
Private Const foo As String = ""Something"", bar As Long = 42
";
            const string expected = @"
Private Const foo As String = ""Something""
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(content, out _).Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                if (state.Status != ParserState.Ready)
                {
                    Assert.Inconclusive("Parser isn't ready. Test cannot proceed.");
                }

                var declarations = state.AllUserDeclarations;
                var target = declarations.SingleOrDefault(d => d.IdentifierName == "bar" && d.DeclarationType == DeclarationType.Constant);
                if (target == null)
                {
                    Assert.Inconclusive("Target constant was not found in test code.");
                }

                var rewriter = state.GetRewriter(target);
                rewriter.Remove(target);

                Assert.AreEqual(expected, rewriter.GetText());
            }
        }

        [Test]
        [Category("TokenStreamRewriter")]
        public void RemovesMiddleVariableInDeclarationList()
        {
            const string content = @"
Private foo As String, bar As Long, buzz As Integer
";
            const string expected = @"
Private foo As String, buzz As Integer
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(content, out _).Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                if (state.Status != ParserState.Ready)
                {
                    Assert.Inconclusive("Parser isn't ready. Test cannot proceed.");
                }

                var declarations = state.AllUserDeclarations;
                var target = declarations.SingleOrDefault(d => d.IdentifierName == "bar" && d.DeclarationType == DeclarationType.Variable);
                if (target == null)
                {
                    Assert.Inconclusive("Target variable was not found in test code.");
                }

                var rewriter = state.GetRewriter(target);
                rewriter.Remove(target);

                Assert.AreEqual(expected, rewriter.GetText());
            }
        }

        [Test]
        [Category("TokenStreamRewriter")]
        public void RemovesMiddleConstantInDeclarationList()
        {
            const string content = @"
Private Const foo As String = ""Something"", bar As Long = 42, buzz As Integer = 12
";
            const string expected = @"
Private Const foo As String = ""Something"", buzz As Integer = 12
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(content, out _).Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                if (state.Status != ParserState.Ready)
                {
                    Assert.Inconclusive("Parser isn't ready. Test cannot proceed.");
                }

                var declarations = state.AllUserDeclarations;
                var target = declarations.SingleOrDefault(d => d.IdentifierName == "bar" && d.DeclarationType == DeclarationType.Constant);
                if (target == null)
                {
                    Assert.Inconclusive("Target constant was not found in test code.");
                }

                var rewriter = state.GetRewriter(target);
                rewriter.Remove(target);

                Assert.AreEqual(expected, rewriter.GetText());
            }
        }

        [Test]
        [Category("TokenStreamRewriter")]
        public void RemovesMiddleVariableInDeclarationListWithLineContinuations()
        {
            const string content = @"
Private foo As String, _
        bar As Long, _
        buzz As Integer
";
            const string expected = @"
Private foo As String, _
        buzz As Integer
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(content, out _).Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                if (state.Status != ParserState.Ready)
                {
                    Assert.Inconclusive("Parser isn't ready. Test cannot proceed.");
                }

                var declarations = state.AllUserDeclarations;
                var target = declarations.SingleOrDefault(d => d.IdentifierName == "bar" && d.DeclarationType == DeclarationType.Variable);
                if (target == null)
                {
                    Assert.Inconclusive("Target variable was found in test code.");
                }

                var rewriter = state.GetRewriter(target);
                rewriter.Remove(target);

                Assert.AreEqual(expected, rewriter.GetText());
            }
        }

        [Test]
        [Category("TokenStreamRewriter")]
        public void RemovesMiddleConstantInDeclarationListWithLineContinuations()
        {
            const string content = @"
Private Const foo _
          As String = ""Something"", _
        bar As Long _
          = 42, _
        buzz As Integer = 12
";
            const string expected = @"
Private Const foo _
          As String = ""Something"", _
        buzz As Integer = 12
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(content, out _).Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                if (state.Status != ParserState.Ready)
                {
                    Assert.Inconclusive("Parser isn't ready. Test cannot proceed.");
                }

                var declarations = state.AllUserDeclarations;
                var target = declarations.SingleOrDefault(d => d.IdentifierName == "bar" && d.DeclarationType == DeclarationType.Constant);
                if (target == null)
                {
                    Assert.Inconclusive("Target constant was found in test code.");
                }

                var rewriter = state.GetRewriter(target);
                rewriter.Remove(target);

                Assert.AreEqual(expected, rewriter.GetText());
            }
        }

        private IProjectsProvider TestProvider(QualifiedModuleName module, ICodeModule testModule)
        {
            var component = new Mock<IVBComponent>();
            component.Setup(c => c.CodeModule).Returns(testModule);
            var provider = new Mock<IProjectsProvider>();
            provider.Setup(p => p.Component(It.IsAny<QualifiedModuleName>()))
                .Returns<QualifiedModuleName>(qmn => qmn.Equals(module) ? component.Object : null);
            return provider.Object;
        }
    }
}
