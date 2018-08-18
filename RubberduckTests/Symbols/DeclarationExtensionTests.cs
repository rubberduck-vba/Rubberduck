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

Public Property Get TestInterface_Foo(Bar As Long) As Long
End Property

Public Property Let TestInterface_Foo(Bar As Long, RHS As Long)
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

        [Test]
        [Category("Resolver")]
        public void FindInterfaceMembersMatchesPublicVariables()
        {
            var intrface =
                @"Option Explicit

Public Bar As String

Public Property Get Foo() As Long
End Property

Public Property Let Foo(rhs As Long)
End Property
";

            var implementation =
                @"Option Explicit

Implements TestInterface

Public Property Get TestInterface_Bar() As String
End Property

Public Property Let TestInterface_Bar(rhs As String)
End Property

Public Property Get TestInterface_Foo() As Long
End Property

Public Property Let TestInterface_Foo(rhs As Long)
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
            var actual = declarations.FindInterfaceMembers().ToList();
            var results = actual.Count;
            var expected = 3;

            Assert.AreEqual(expected, results, "Expected {0} Declarations, received {1}", expected, results);
        }

        [Test]
        [Category("Resolver")]
        public void FindInterfaceImplementationMembersMatchesPublicVariables()
        {
            var intrface =
                @"Option Explicit

Public Bar As String

Public Property Get Foo() As Long
End Property

Public Property Let Foo(rhs As Long)
End Property
";

            var implementation =
                @"Option Explicit

Implements TestInterface

Public Property Get TestInterface_Bar() As String
End Property

Public Property Let TestInterface_Bar(rhs As String)
End Property

Public Property Get TestInterface_Foo() As Long
End Property

Public Property Let TestInterface_Foo(rhs As Long)
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
            var declaration = declarations.Single(decl => decl.DeclarationType == DeclarationType.Variable && decl.IdentifierName.Equals("Bar"));

            var expected = declarations.Where(decl => decl.DeclarationType.HasFlag(DeclarationType.Property) && decl.IdentifierName.Equals("TestInterface_Bar")).ToList();
            var actual = declarations.FindInterfaceImplementationMembers(declaration).ToList();
            var results = actual.Count;

            Assert.AreEqual(expected.Count, results, "Expected {0} Declarations, received {1}", expected.Count, results);
            Assert.That(actual, Is.EquivalentTo(expected));
        }

        [Test]
        [Category("Resolver")]
        public void FindInterfaceImplementationMembersPublicVariantMatchesAllPropertyTypes()
        {
            var intrface =
                @"Option Explicit

Public Bar As Variant";

            var implementation =
                @"Option Explicit

Implements TestInterface

Public Property Get TestInterface_Bar() As Variant
End Property

Public Property Let TestInterface_Bar(rhs As Variant)
End Property

Public Property Set TestInterface_Bar(rhs As Variant)
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
            var declaration = declarations.Single(decl => decl.DeclarationType == DeclarationType.Variable && decl.IdentifierName.Equals("Bar"));

            var expected = declarations.Where(decl => decl.DeclarationType.HasFlag(DeclarationType.Property) && decl.IdentifierName.Equals("TestInterface_Bar")).ToList();
            var actual = declarations.FindInterfaceImplementationMembers(declaration).ToList();
            var results = actual.Count;

            Assert.AreEqual(expected.Count, results, "Expected {0} Declarations, received {1}", expected.Count, results);
            Assert.That(actual, Is.EquivalentTo(expected));
        }

        [Test]
        [Category("Resolver")]
        public void FindInterfaceImplementationMembersPublicIntrinsicDoesNotMatchSet()
        {
            var intrface =
                @"Option Explicit

Public Bar As Long";

            var implementation =
                @"Option Explicit

Implements TestInterface

Public Property Get TestInterface_Bar() As Long
End Property

Public Property Let TestInterface_Bar(rhs As Long)
End Property

Public Property Set TestInterface_Bar(rhs As Variant)
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
            var declaration = declarations.Single(decl => decl.DeclarationType == DeclarationType.Variable && decl.IdentifierName.Equals("Bar"));

            var expected = declarations.Where( decl => decl.IdentifierName.Equals("TestInterface_Bar") &&
                        decl.DeclarationType == DeclarationType.PropertyLet ||
                        decl.DeclarationType == DeclarationType.PropertyGet).ToList();
            var actual = declarations.FindInterfaceImplementationMembers(declaration).ToList();
            var results = actual.Count;

            Assert.AreEqual(expected.Count, results, "Expected {0} Declarations, received {1}", expected.Count, results);
            Assert.That(actual, Is.EquivalentTo(expected));
        }

        [Test]
        [Category("Resolver")]
        public void FindInterfaceImplementationMembersPublicObjectDoesNotMatchLet()
        {
            var intrface =
                @"Option Explicit

Public Bar As Object";

            var implementation =
                @"Option Explicit

Implements TestInterface

Public Property Get TestInterface_Bar() As Object
End Property

Public Property Let TestInterface_Bar(rhs As Variant)
End Property

Public Property Set TestInterface_Bar(rhs As Object)
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
            var declaration = declarations.Single(decl => decl.DeclarationType == DeclarationType.Variable && decl.IdentifierName.Equals("Bar"));

            var expected = declarations.Where(decl => decl.IdentifierName.Equals("TestInterface_Bar") &&
                                                      decl.DeclarationType == DeclarationType.PropertySet ||
                                                      decl.DeclarationType == DeclarationType.PropertyGet).ToList();
            var actual = declarations.FindInterfaceImplementationMembers(declaration).ToList();
            var results = actual.Count;

            Assert.AreEqual(expected.Count, results, "Expected {0} Declarations, received {1}", expected.Count, results);
            Assert.That(actual, Is.EquivalentTo(expected));
        }

        [Test]
        [Category("Resolver")]
        public void FindFindInterfaceMemberMatchesDeclarationTypes()
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

Public Property Get TestInterface_Foo(Bar As Long) As Long
End Property

Public Property Let TestInterface_Foo(Bar As Long, RHS As Long)
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
            var declaration = declarations.Single(decl => decl.DeclarationType == DeclarationType.PropertyGet && decl.IdentifierName.Equals("TestInterface_Foo"));

            var expected = declarations.Single(decl => decl.DeclarationType == DeclarationType.PropertyGet && decl.IdentifierName.Equals("Foo"));
            var actual = declarations.FindInterfaceMember(declaration);

            Assert.AreEqual(expected, actual, "Expected {0}, resolved to {1}", expected, actual);
        }

        [Test]
        [Category("Resolver")]
        public void FindFindInterfaceMemberNoResultWithoutMatchingDeclaration()
        {
            var intrface =
                @"Option Explicit

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
                .AddComponent("TestInterface", ComponentType.ClassModule, intrface, new Selection(1, 1))
                .AddComponent("TestImplementation", ComponentType.ClassModule, implementation, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            var parser = MockParser.Create(vbe.Object);
            parser.Parse(new CancellationTokenSource());

            var declarations = parser.State.DeclarationFinder.AllDeclarations.ToList();
            var declaration = declarations.Single(decl => decl.DeclarationType == DeclarationType.PropertyGet && decl.IdentifierName.Equals("TestInterface_Foo"));

            var actual = declarations.FindInterfaceMember(declaration);

            Assert.IsNull(actual, "Expected null, resolved to {0}", actual);
        }

        [Test]
        [Category("Resolver")]
        public void ImplementsInterfaceMemberMatchesProperty()
        {
            var intrface =
                @"Option Explicit

Public Property Get Foo() As Long
End Property

Public Property Let Foo(rhs As Long)
End Property
";

            var implementation =
                @"Option Explicit

Implements TestInterface

Public Property Get TestInterface_Foo() As Long
End Property

Public Property Let TestInterface_Foo(rhs As Long)
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
            var member = declarations.Single(decl => decl.IdentifierName.Equals("Foo") && decl.DeclarationType == DeclarationType.PropertyGet);
            var implementing = declarations.Single(decl => decl.IdentifierName.Equals("TestInterface_Foo") && decl.DeclarationType == DeclarationType.PropertyGet);

            Assert.IsTrue(implementing.ImplementsInterfaceMember(member));
        }

        [Test]
        [Category("Resolver")]
        public void ImplementsInterfaceMemberMatchesPublicVariable()
        {
            var intrface =
                @"Option Explicit

Public Foo As Long";

            var implementation =
                @"Option Explicit

Implements TestInterface

Public Property Get TestInterface_Foo() As Long
End Property

Public Property Let TestInterface_Foo(rhs As Long)
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
            var member = declarations.Single(decl => decl.IdentifierName.Equals("Foo") && decl.DeclarationType == DeclarationType.Variable);
            var implementing = declarations.Single(decl => decl.IdentifierName.Equals("TestInterface_Foo") && decl.DeclarationType == DeclarationType.PropertyGet);

            Assert.IsTrue(implementing.ImplementsInterfaceMember(member));
        }

        [Test]
        [Category("Resolver")]
        public void ImplementsInterfaceMemberLetMatchesPublicIntrinsic()
        {
            var intrface =
                @"Option Explicit

Public Foo As Long";

            var implementation =
                @"Option Explicit

Implements TestInterface

Public Property Get TestInterface_Foo() As Long
End Property

Public Property Let TestInterface_Foo(rhs As Long)
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
            var member = declarations.Single(decl => decl.IdentifierName.Equals("Foo") && decl.DeclarationType == DeclarationType.Variable);
            var implementing = declarations.Single(decl => decl.IdentifierName.Equals("TestInterface_Foo") && decl.DeclarationType == DeclarationType.PropertyLet);

            Assert.IsTrue(implementing.ImplementsInterfaceMember(member));
        }

        [Test]
        [Category("Resolver")]
        public void ImplementsInterfaceMemberSetDoesNotMatchPublicIntrinsic()
        {
            var intrface =
                @"Option Explicit

Public Foo As Long";

            var implementation =
                @"Option Explicit

Implements TestInterface

Public Property Get TestInterface_Foo() As Long
End Property

Public Property Set TestInterface_Foo(rhs As Variant)
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
            var member = declarations.Single(decl => decl.IdentifierName.Equals("Foo") && decl.DeclarationType == DeclarationType.Variable);
            var implementing = declarations.Single(decl => decl.IdentifierName.Equals("TestInterface_Foo") && decl.DeclarationType == DeclarationType.PropertySet);

            Assert.IsFalse(implementing.ImplementsInterfaceMember(member));
        }

        [Test]
        [Category("Resolver")]
        public void ImplementsInterfaceMemberSetMatchesPublicObject()
        {
            var intrface =
                @"Option Explicit

Public Foo As Object";

            var implementation =
                @"Option Explicit

Implements TestInterface

Public Property Get TestInterface_Foo() As Object
End Property

Public Property Set TestInterface_Foo(rhs As Object)
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
            var member = declarations.Single(decl => decl.IdentifierName.Equals("Foo") && decl.DeclarationType == DeclarationType.Variable);
            var implementing = declarations.Single(decl => decl.IdentifierName.Equals("TestInterface_Foo") && decl.DeclarationType == DeclarationType.PropertySet);

            Assert.IsTrue(implementing.ImplementsInterfaceMember(member));
        }

        [Test]
        [Category("Resolver")]
        public void ImplementsInterfaceMemberLetDoesNotMatchPublicObject()
        {
            var intrface =
                @"Option Explicit

Public Foo As Object";

            var implementation =
                @"Option Explicit

Implements TestInterface

Public Property Get TestInterface_Foo() As Object
End Property

Public Property Let TestInterface_Foo(rhs As Variant)
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
            var member = declarations.Single(decl => decl.IdentifierName.Equals("Foo") && decl.DeclarationType == DeclarationType.Variable);
            var implementing = declarations.Single(decl => decl.IdentifierName.Equals("TestInterface_Foo") && decl.DeclarationType == DeclarationType.PropertyLet);

            Assert.IsFalse(implementing.ImplementsInterfaceMember(member));
        }

        [Test]
        [Category("Resolver")]
        public void ImplementsInterfaceMemberVariantMatchesLetAndSet()
        {
            var intrface =
                @"Option Explicit

Public Foo As Variant";

            var implementation =
                @"Option Explicit

Implements TestInterface

Public Property Get TestInterface_Foo() As Variant
End Property

Public Property Let TestInterface_Foo(rhs As Variant)
End Property

Public Property Set TestInterface_Foo(rhs As Variant)
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
            var member = declarations.Single(decl => decl.IdentifierName.Equals("Foo") && decl.DeclarationType == DeclarationType.Variable);
            var setter = declarations.Single(decl => decl.IdentifierName.Equals("TestInterface_Foo") && decl.DeclarationType == DeclarationType.PropertySet);
            var letter = declarations.Single(decl => decl.IdentifierName.Equals("TestInterface_Foo") && decl.DeclarationType == DeclarationType.PropertyLet);

            Assert.IsTrue(setter.ImplementsInterfaceMember(member));
            Assert.IsTrue(letter.ImplementsInterfaceMember(member));
        }
    }
}
