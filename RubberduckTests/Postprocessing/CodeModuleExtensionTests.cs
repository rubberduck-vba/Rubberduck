using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Antlr4.Runtime;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

// ReSharper disable InvokeAsExtensionMethod
namespace RubberduckTests.Postprocessing
{
    [TestClass]
    public class CodeModuleExtensionTests
    {
        [TestMethod]
        [TestCategory("TokenStreamRewriter")]
        public void RewriteClearsEntireModule()
        {
            var module = new Mock<ICodeModule>();
            module.Setup(m => m.Clear());

            var rewriter = new TokenStreamRewriter(new CommonTokenStream(new ListTokenSource(new List<IToken>())));

            CodeModuleExtensions.Rewrite(module.Object, rewriter);
            module.Verify(m => m.Clear());
        }

        [TestMethod]
        [TestCategory("TokenStreamRewriter")]
        public void RewriterInsertsRewriterOutputAtLine1()
        {
            const string content = @"Option Explicit";
            IVBComponent component;
            var vbe = new MockVbeBuilder().BuildFromSingleStandardModule(content, out component).Object;
            var module = component.CodeModule;

            var parser = MockParser.Create(vbe, new RubberduckParserState(vbe));
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status != ParserState.Ready)
            {
                Assert.Inconclusive("Parser isn't ready. Test cannot proceed.");
            }

            var rewriter = parser.State.GetRewriter(component);

            CodeModuleExtensions.Rewrite(module, rewriter);
            Assert.AreEqual(content, rewriter.GetText());
        }

        [TestMethod]
        [TestCategory("TokenStreamRewriter")]
        public void RemovesSimpleVariableDeclaration()
        {
            const string expected = @"
";
            const string content = @"
Private foo As String
";
            IVBComponent component;
            var vbe = new MockVbeBuilder().BuildFromSingleStandardModule(content, out component).Object;
            var parser = MockParser.Create(vbe, new RubberduckParserState(vbe));
            
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status != ParserState.Ready)
            {
                Assert.Inconclusive("Parser isn't ready. Test cannot proceed.");
            }

            var declarations = parser.State.AllUserDeclarations;
            var target = declarations.SingleOrDefault(d => d.DeclarationType == DeclarationType.Variable);
            if (target == null)
            {
                Assert.Inconclusive("No variable was found in test code.");
            }

            var module = component.CodeModule;
            var rewriter = parser.State.GetRewriter(target);
            
            CodeModuleExtensions.Remove(module, rewriter, target);
            Assert.AreEqual(expected, rewriter.GetText());
        }

        [TestMethod]
        [TestCategory("TokenStreamRewriter")]
        public void RemovesSimpleVariableStatement()
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
            IVBComponent component;
            var vbe = new MockVbeBuilder().BuildFromSingleStandardModule(content, out component).Object;
            var parser = MockParser.Create(vbe, new RubberduckParserState(vbe));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status != ParserState.Ready)
            {
                Assert.Inconclusive("Parser isn't ready. Test cannot proceed.");
            }

            var declarations = parser.State.AllUserDeclarations;
            var target = declarations.SingleOrDefault(d => d.DeclarationType == DeclarationType.Variable);
            if (target == null)
            {
                Assert.Inconclusive("No variable was found in test code.");
            }

            var module = component.CodeModule;
            var rewriter = parser.State.GetRewriter(target);

            CodeModuleExtensions.Remove(module, rewriter, target);
            Assert.AreEqual(expected, rewriter.GetText());
        }

        [TestMethod]
        [TestCategory("TokenStreamRewriter")]
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
            IVBComponent component;
            var vbe = new MockVbeBuilder().BuildFromSingleStandardModule(content, out component).Object;
            var parser = MockParser.Create(vbe, new RubberduckParserState(vbe));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status != ParserState.Ready)
            {
                Assert.Inconclusive("Parser isn't ready. Test cannot proceed.");
            }

            var declarations = parser.State.AllUserDeclarations;
            var target = declarations.SingleOrDefault(d => d.IdentifierName == "foo" && d.DeclarationType == DeclarationType.Variable);
            if (target == null)
            {
                Assert.Inconclusive("No variable was found in test code.");
            }

            var module = component.CodeModule;
            var rewriter = parser.State.GetRewriter(target);

            CodeModuleExtensions.Remove(module, rewriter, target);
            Assert.AreEqual(expected, rewriter.GetText());
        }

        [TestMethod]
        [TestCategory("TokenStreamRewriter")]
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
            IVBComponent component;
            var vbe = new MockVbeBuilder().BuildFromSingleStandardModule(content, out component).Object;
            var parser = MockParser.Create(vbe, new RubberduckParserState(vbe));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status != ParserState.Ready)
            {
                Assert.Inconclusive("Parser isn't ready. Test cannot proceed.");
            }

            var declarations = parser.State.AllUserDeclarations;
            var target = declarations.SingleOrDefault(d => d.IdentifierName == "bar" && d.DeclarationType == DeclarationType.Variable);
            if (target == null)
            {
                Assert.Inconclusive("No variable was found in test code.");
            }

            var module = component.CodeModule;
            var rewriter = parser.State.GetRewriter(target);

            CodeModuleExtensions.Remove(module, rewriter, target);
            Assert.AreEqual(expected, rewriter.GetText());
        }

        [TestMethod]
        [TestCategory("TokenStreamRewriter")]
        public void RemovesSingleVariableDeclarationWithLineContinuations()
        {
            const string expected = @"
";
            const string content = @"
Private foo _
  As _
    String
";
            IVBComponent component;
            var vbe = new MockVbeBuilder().BuildFromSingleStandardModule(content, out component).Object;
            var parser = MockParser.Create(vbe, new RubberduckParserState(vbe));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status != ParserState.Ready)
            {
                Assert.Inconclusive("Parser isn't ready. Test cannot proceed.");
            }

            var declarations = parser.State.AllUserDeclarations;
            var target = declarations.SingleOrDefault(d => d.DeclarationType == DeclarationType.Variable);
            if (target == null)
            {
                Assert.Inconclusive("No variable was found in test code.");
            }

            var module = component.CodeModule;
            var rewriter = parser.State.GetRewriter(target);

            CodeModuleExtensions.Remove(module, rewriter, target);
            Assert.AreEqual(expected, rewriter.GetText());
        }

        [TestMethod]
        [TestCategory("TokenStreamRewriter")]
        public void RemovesFirstVariableInDeclarationList()
        {
            const string content = @"
Private foo As String, bar As Long
";
            const string expected = @"
Private bar As Long
";
            IVBComponent component;
            var vbe = new MockVbeBuilder().BuildFromSingleStandardModule(content, out component).Object;
            var parser = MockParser.Create(vbe, new RubberduckParserState(vbe));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status != ParserState.Ready)
            {
                Assert.Inconclusive("Parser isn't ready. Test cannot proceed.");
            }

            var declarations = parser.State.AllUserDeclarations;
            var target = declarations.SingleOrDefault(d => d.IdentifierName == "foo" && d.DeclarationType == DeclarationType.Variable);
            if (target == null)
            {
                Assert.Inconclusive("Target variable was found in test code.");
            }

            var module = component.CodeModule;
            var rewriter = parser.State.GetRewriter(target);

            CodeModuleExtensions.Remove(module, rewriter, target);
            Assert.AreEqual(expected, rewriter.GetText());
        }

        [TestMethod]
        [TestCategory("TokenStreamRewriter")]
        public void RemovesLastVariableInDeclarationList()
        {
            const string content = @"
Private foo As String, bar As Long
";
            const string expected = @"
Private foo As String
";
            IVBComponent component;
            var vbe = new MockVbeBuilder().BuildFromSingleStandardModule(content, out component).Object;
            var parser = MockParser.Create(vbe, new RubberduckParserState(vbe));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status != ParserState.Ready)
            {
                Assert.Inconclusive("Parser isn't ready. Test cannot proceed.");
            }

            var declarations = parser.State.AllUserDeclarations;
            var target = declarations.SingleOrDefault(d => d.IdentifierName == "bar" && d.DeclarationType == DeclarationType.Variable);
            if (target == null)
            {
                Assert.Inconclusive("Target variable was found in test code.");
            }

            var module = component.CodeModule;
            var rewriter = parser.State.GetRewriter(target);

            CodeModuleExtensions.Remove(module, rewriter, target);
            Assert.AreEqual(expected, rewriter.GetText());
        }

        [TestMethod]
        [TestCategory("TokenStreamRewriter")]
        public void RemovesMiddleVariableInDeclarationList()
        {
            const string content = @"
Private foo As String, bar As Long, buzz As Integer
";
            const string expected = @"
Private foo As String, buzz As Integer
";
            IVBComponent component;
            var vbe = new MockVbeBuilder().BuildFromSingleStandardModule(content, out component).Object;
            var parser = MockParser.Create(vbe, new RubberduckParserState(vbe));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status != ParserState.Ready)
            {
                Assert.Inconclusive("Parser isn't ready. Test cannot proceed.");
            }

            var declarations = parser.State.AllUserDeclarations;
            var target = declarations.SingleOrDefault(d => d.IdentifierName == "bar" && d.DeclarationType == DeclarationType.Variable);
            if (target == null)
            {
                Assert.Inconclusive("Target variable was found in test code.");
            }

            var module = component.CodeModule;
            var rewriter = parser.State.GetRewriter(target);

            CodeModuleExtensions.Remove(module, rewriter, target);
            Assert.AreEqual(expected, rewriter.GetText());
        }

        [TestMethod]
        [TestCategory("TokenStreamRewriter")]
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
            IVBComponent component;
            var vbe = new MockVbeBuilder().BuildFromSingleStandardModule(content, out component).Object;
            var parser = MockParser.Create(vbe, new RubberduckParserState(vbe));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status != ParserState.Ready)
            {
                Assert.Inconclusive("Parser isn't ready. Test cannot proceed.");
            }

            var declarations = parser.State.AllUserDeclarations;
            var target = declarations.SingleOrDefault(d => d.IdentifierName == "bar" && d.DeclarationType == DeclarationType.Variable);
            if (target == null)
            {
                Assert.Inconclusive("Target variable was found in test code.");
            }

            var module = component.CodeModule;
            var rewriter = parser.State.GetRewriter(target);

            CodeModuleExtensions.Remove(module, rewriter, target);
            Assert.AreEqual(expected, rewriter.GetText());
        }
    }
}
