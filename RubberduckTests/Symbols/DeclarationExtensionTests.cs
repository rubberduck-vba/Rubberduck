using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Symbols
{
    [TestFixture]
    public class DeclarationExtensionTests
    {
        [Test]
        [Category("Resolver")]
        public void FindInterfaceImplementationMembersMatchesDeclarationTypes()
        {
            var intrface =
@"Option Explicit

Public Property Get Foo(Bar As Long) As Long
End Property

Public Property Let Foo(Bar As Long, NewValue As Long)
End Property
";

            var implementation =
@"Option Explicit

Implements TestInterface

Private Property Get TestInterface_Foo(Bar As Long) As Long
End Property

Private Property Let TestInterface_Foo(Bar As Long, RHS As Long)
End Property
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("UnderTest", ProjectProtection.Unprotected)
                .AddComponent("TestInterface", ComponentType.ClassModule, intrface, new Selection(1, 1))
                .AddComponent("TestImplementation", ComponentType.ClassModule, implementation, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            var parser = MockParser.Create(vbe.Object);
            parser.Parse(new CancellationTokenSource());

            var declarations = parser.State.DeclarationFinder.AllDeclarations.ToList();
            var declaration = declarations.Single(decl => decl.DeclarationType == DeclarationType.PropertyGet && decl.IdentifierName.Equals("Foo"));

            var expected = declarations.Single(decl => decl.DeclarationType == DeclarationType.PropertyGet && decl.IdentifierName.Equals("TestInterface_Foo"));
            var actual = declarations.FindInterfaceImplementationMembers(declaration).ToList();
            var results = actual.Count;

            Assert.AreEqual(1, results, "Expected {0} Declarations, received {1}", expected, results);
            Assert.AreEqual(expected, actual.First(), "Expected {0}, resolved to {1}", expected, actual);
        }
    }
}
