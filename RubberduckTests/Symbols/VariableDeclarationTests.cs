using System.Collections.Generic;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Symbols
{
    [TestFixture]
    public class VariableDeclarationTests
    {
        [Test]
        [TestCase("Private WithEvents foo As EventSource", true)]
        [TestCase("Private foo As EventSource", false)]
        [Category("Resolver")]
        public void WithEventsIsResolvedCorrectly(string declaration, bool withEvents)
        {
            var variables = ArrangeAndGetVariableDeclarations(ComponentType.ClassModule, declaration);
            var foo = variables.Single();
            Assert.AreEqual(withEvents, foo.IsWithEvents);  
        }

        [Test]
        [TestCase("Private WithEvents {0} As EventSource, WithEvents {1} As EventSource", "foo", "bar", true, true)]
        [TestCase("Private {0} As EventSource, WithEvents {1} As EventSource", "foo", "bar", false, true)]
        [TestCase("Private WithEvents {0} As EventSource, {1} As EventSource", "foo", "bar", true, false)]
        [TestCase("Private {0} As EventSource, {1} As EventSource", "foo", "bar", false, false)]
        [Category("Resolver")]
        public void WithEventsIsResolvedCorrectlyVariableList(string template, string first, string second, bool firstEvents, bool secondEvents)
        {
            var variables = ArrangeAndGetVariableDeclarations(ComponentType.ClassModule, string.Format(template, first, second));
            Assert.AreEqual(2, variables.Count);
            Assert.AreEqual(firstEvents, variables.Single(variable => variable.IdentifierName.Equals(first)).IsWithEvents);
            Assert.AreEqual(secondEvents, variables.Single(variable => variable.IdentifierName.Equals(second)).IsWithEvents);
        }

        private List<VariableDeclaration> ArrangeAndGetVariableDeclarations(ComponentType moduleType, string code)
        {
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("UnderTest", moduleType, code, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            using (var parser = MockParser.Create(vbe.Object))
            {
                parser.Parse(new CancellationTokenSource());

                return parser.State.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                    .Cast<VariableDeclaration>()
                    .ToList();
            }
        } 
    }
}
