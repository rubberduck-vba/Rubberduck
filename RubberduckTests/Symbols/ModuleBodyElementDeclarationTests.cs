using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Symbols
{
    [TestFixture]
    public class ModuleBodyElementDeclarationTests
    {
        [Test]
        [Category("Resolver")]
        public void FindFindInterfaceMemberMatchesDeclarationTypes()
        {
            var interfaceModule =
                @"Option Explicit

Public Property Get Foo(Bar As Long) As Long
End Property

Public Property Let Foo(Bar As Long, NewValue As Long)
End Property
        ";

            var implementation =
                @"Option Explicit

Implements TestInterface

Public Property Get TestInterface_Foo(Bar As Long) As Long
End Property

Public Property Let TestInterface_Foo(Bar As Long, RHS As Long)
End Property
        ";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("UnderTest", ProjectProtection.Unprotected)
                .AddComponent("TestInterface", ComponentType.ClassModule, interfaceModule, new Selection(1, 1))
                .AddComponent("TestImplementation", ComponentType.ClassModule, implementation, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            var parser = MockParser.Create(vbe.Object);
            parser.Parse(new CancellationTokenSource());

            var declarations = parser.State.DeclarationFinder.AllDeclarations.ToList();

            var implementing = (ModuleBodyElementDeclaration) declarations.Single(decl =>
                decl.DeclarationType == DeclarationType.PropertyGet && decl.IdentifierName.Equals("TestInterface_Foo"));
            var member = (ModuleBodyElementDeclaration) declarations.Single(decl =>
                decl.DeclarationType == DeclarationType.PropertyGet && decl.IdentifierName.Equals("Foo"));

            Assert.AreSame(member, implementing.InterfaceMemberImplemented);
        }
    }
}
