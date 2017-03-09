using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Symbols
{
    [TestClass]
    public class DeclarationFinderTests
    {
        [TestMethod]
        [Ignore] // ref. https://github.com/rubberduck-vba/Rubberduck/issues/2330
        public void FiendishlyAmbiguousNameSelectsSmallestScopedDeclaration()
        {
            var code = @"
Option Explicit

Public Sub foo()
    Dim foo As Long
    foo = 42
    Debug.Print foo
End Sub
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("foo", ProjectProtection.Unprotected)
                .AddComponent("foo", ComponentType.StandardModule, code, new Selection(6, 6))
                .MockVbeBuilder()
                .Build();

            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));
            parser.Parse(new CancellationTokenSource());

            var expected = parser.State.AllDeclarations.Single(item => item.DeclarationType == DeclarationType.Variable);
            var actual = parser.State.DeclarationFinder.FindSelectedDeclaration(vbe.Object.ActiveCodePane);

            Assert.AreEqual(expected, actual, "Expected {0}, resolved to {1}", expected.DeclarationType, actual.DeclarationType);
        }

        [TestMethod]
        [Ignore] // bug: this test should pass... it's not all that evil
        public void AmbiguousNameSelectsSmallestScopedDeclaration()
        {
            var code = @"
Option Explicit

Public Sub foo()
    Dim foo As Long
    foo = 42
    Debug.Print foo
End Sub
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("TestModule", ComponentType.StandardModule, code, new Selection(6, 6))
                .MockVbeBuilder()
                .Build();

            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));
            parser.Parse(new CancellationTokenSource());

            var expected = parser.State.AllDeclarations.Single(item => item.DeclarationType == DeclarationType.Variable);
            var actual = parser.State.DeclarationFinder.FindSelectedDeclaration(vbe.Object.ActiveCodePane);

            Assert.AreEqual(expected, actual, "Expected {0}, resolved to {1}", expected.DeclarationType, actual.DeclarationType);
        }
    }
}