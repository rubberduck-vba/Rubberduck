using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Rubberduck.VBEditor;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;

namespace RubberduckTests.Symbols
{

    [TestFixture]
    public class DeclarationFinderTests
    {
        [Test]
        [Category("Resolver")]
        public void SameNameForProjectAndClassImplicit_ScopedDeclaration()
        {
            var refEditClass = @"
Option Explicit

Private ValueField As Variant

Public Property Get Value()
  Value = ValueField
End Property

Public Property Let Value(Value As Variant)
  ValueField = Value
End Property";

            var code =
                @"
Option Explicit

Public Sub foo()
    Dim myEdit As RefEdit
    Set myEdit = New RefEdit

    myEdit.Value = ""abc""
End Sub
";

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("RefEdit", ProjectProtection.Unprotected)
                .AddComponent("RefEdit", ComponentType.ClassModule, refEditClass)
                .AddComponent("Test", ComponentType.StandardModule, code, new Selection(7, 6))
                .AddProjectToVbeBuilder()
                .Build();

            var parser = MockParser.Create(vbe.Object);
            parser.Parse(new CancellationTokenSource());

            var expected = ParserState.ResolverError;
            var actual = parser.State.Status;

            Assert.AreEqual(expected, actual, "Expected {0}, resolved to {1}", expected, actual);
        }

        [Category("Resolver")]
        [Test]
        public void Identify_NamedParameter_Parameter_FromExcel_DefaultAccess()
        {
            // Note that ColumnIndex is actually a parameter of the _Default default member
            // of the Excel.Range object.
            const string code = @"
Public Sub DoIt()
    Dim foo As Variant
    Dim sht As WorkSheet

    foo = sht.Cells(ColumnIndex:=12).Value
End Sub";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("TestModule", ComponentType.StandardModule, code)
                .AddReference(ReferenceLibrary.Excel)
                .AddProjectToVbeBuilder()
                .Build();

            var selection = new Selection(6, 21, 6, 32);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var module = state.DeclarationFinder.AllModules.First(qmn => qmn.ComponentName.Equals("TestModule"));
                var qualifiedSelection = new QualifiedSelection(module, selection);

                var reference = state.DeclarationFinder.IdentifierReferences(qualifiedSelection).First();
                var referencedDeclaration = reference.Declaration;

                var expectedReferencedDeclarationName = "EXCEL.EXE;Excel.Range._Default.Let.ColumnIndex";
                var actualReferencedDeclarationName = $"{referencedDeclaration.ParentScope}.{referencedDeclaration.IdentifierName}";

                Assert.AreEqual(expectedReferencedDeclarationName, actualReferencedDeclarationName);
                Assert.AreEqual(DeclarationType.Parameter, referencedDeclaration.DeclarationType);
            }
        }

        [Test]
        [Category("Resolver")]
        public void FindParameterFromArgument_WorksWithMultipleScopes()
        {
            var module1 =
@"Public Sub Foo(arg As Variant)
End Sub";

            var module2 =
@"Private Sub Foo(expected As Variant)
End Sub

Public Sub Bar()
    Dim fooBar As Variant
    Foo fooBar
End Sub
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("UnderTest", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, module1, new Selection(1, 1))
                .AddComponent("Module2", ComponentType.StandardModule, module2, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var expected = declarations.FirstOrDefault(decl => decl.IdentifierName.Equals("expected"));

                var enclosing = declarations.FirstOrDefault(decl => decl.IdentifierName.Equals("Bar"));
                var context = enclosing?.Context.GetDescendent<VBAParser.ArgumentExpressionContext>();
                var actual = state.DeclarationFinder.FindParameterOfNonDefaultMemberFromSimpleArgumentNotPassedByValExplicitly(context, enclosing);

                Assert.AreEqual(expected, actual);
            }
        }


        [Category("Resolver")]
        [Category("Interfaces")]
        [Test]
        public void DeclarationFinderCanCopeWithMultipleModulesImplementingTheSameInterface()
        {
            const string interfaceCode = @"
Public Sub Foo()
End Sub
";

            const string implementationCode = @"
Implements IClass1

Public Sub IClass1_Foo()
End Sub
";

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, interfaceCode, new Selection(0, 0))
                .AddComponent("Class2", ComponentType.ClassModule, implementationCode, new Selection(0, 0))
                .AddComponent("Class3", ComponentType.ClassModule, implementationCode, new Selection(0, 0))
                .AddProjectToVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var interfaceDeclarations = state.DeclarationFinder.FindAllInterfaceMembers().ToList();

                Assert.AreEqual(1, interfaceDeclarations.Count());
            }
        }

        [Test]
        [Category("Resolver")]
        [Category("Interfaces")]
        public void FindInterfaceImplementationMembersForGetMatchesOnlyGet()
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

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var declaration = declarations.Single(decl => decl.DeclarationType == DeclarationType.PropertyGet && decl.IdentifierName.Equals("Foo"));

                var expected = declarations.Single(decl => decl.DeclarationType == DeclarationType.PropertyGet && decl.IdentifierName.Equals("TestInterface_Foo"));
                var actual = state.DeclarationFinder.FindInterfaceImplementationMembers(declaration).ToList();
                var results = actual.Count;

                Assert.AreEqual(1, results, "Expected {0} Declarations, received {1}", expected, results);
                Assert.AreEqual(expected, actual.First(), "Expected {0}, resolved to {1}", expected, actual);
            }
        }

        [Test]
        [Category("Resolver")]
        [Category("Interfaces")]
        public void FindInterfaceImplementationMembersForLetMatchesOnlyLet()
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

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var declaration = declarations.Single(decl => decl.DeclarationType == DeclarationType.PropertyLet && decl.IdentifierName.Equals("Foo"));

                var expected = declarations.Single(decl => decl.DeclarationType == DeclarationType.PropertyLet && decl.IdentifierName.Equals("TestInterface_Foo"));
                var actual = state.DeclarationFinder.FindInterfaceImplementationMembers(declaration).ToList();
                var results = actual.Count;

                Assert.AreEqual(1, results, "Expected {0} Declarations, received {1}", expected, results);
                Assert.AreEqual(expected, actual.First(), "Expected {0}, resolved to {1}", expected, actual);
            }
        }

        [Test]
        [Category("Resolver")]
        [Category("Interfaces")]
        public void FindInterfaceImplementationMembersForSetMatchesOnlySet()
        {
            var intrface =
                @"Option Explicit

Public Property Get Foo(Bar As Long) As Object
End Property

Public Property Set Foo(Bar As Long, NewValue As Object)
End Property
";

            var implementation =
                @"Option Explicit

Implements TestInterface

Public Property Get TestInterface_Foo(Bar As Long) As Object
End Property

Public Property Set TestInterface_Foo(Bar As Long, RHS As Object)
End Property
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("UnderTest", ProjectProtection.Unprotected)
                .AddComponent("TestInterface", ComponentType.ClassModule, intrface, new Selection(1, 1))
                .AddComponent("TestImplementation", ComponentType.ClassModule, implementation, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var declaration = declarations.Single(decl => decl.DeclarationType == DeclarationType.PropertySet && decl.IdentifierName.Equals("Foo"));

                var expected = declarations.Single(decl => decl.DeclarationType == DeclarationType.PropertySet && decl.IdentifierName.Equals("TestInterface_Foo"));
                var actual = state.DeclarationFinder.FindInterfaceImplementationMembers(declaration).ToList();
                var results = actual.Count;

                Assert.AreEqual(1, results, "Expected {0} Declarations, received {1}", expected, results);
                Assert.AreEqual(expected, actual.First(), "Expected {0}, resolved to {1}", expected, actual);
            }
        }

        [Test]
        [Category("Resolver")]
        [Category("Interfaces")]
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

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var actual = state.DeclarationFinder.FindAllInterfaceMembers().ToList();
                var expected = state.DeclarationFinder.AllUserDeclarations.Where(decl => decl.ParentScope.Equals("UnderTest.TestInterface")).ToList();

                var results = actual.Count;

                Assert.AreEqual(expected.Count, results, "Expected {0} Declarations, received {1}", expected.Count, results);
                Assert.That(actual, Is.EquivalentTo(expected));
            }
        }

        [Test]
        [Category("Resolver")]
        [Category("Interfaces")]
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

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var declaration = declarations.Single(decl => decl.DeclarationType == DeclarationType.Variable && decl.IdentifierName.Equals("Bar"));

                var expected = declarations.Where(decl => decl.DeclarationType.HasFlag(DeclarationType.Property) && decl.IdentifierName.Equals("TestInterface_Bar")).ToList();
                var actual = state.DeclarationFinder.FindInterfaceImplementationMembers(declaration).ToList();
                var results = actual.Count;

                Assert.AreEqual(expected.Count, results, "Expected {0} Declarations, received {1}", expected.Count, results);
                Assert.That(actual, Is.EquivalentTo(expected));
            }
        }

        [Test]
        [Category("Resolver")]
        [Category("Interfaces")]
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

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var declaration = declarations.Single(decl => decl.DeclarationType == DeclarationType.Variable && decl.IdentifierName.Equals("Bar"));

                var expected = declarations.Where(decl => decl.DeclarationType.HasFlag(DeclarationType.Property) && decl.IdentifierName.Equals("TestInterface_Bar")).ToList();
                var actual = state.DeclarationFinder.FindInterfaceImplementationMembers(declaration).ToList();
                var results = actual.Count;

                Assert.AreEqual(expected.Count, results, "Expected {0} Declarations, received {1}", expected.Count, results);
                Assert.That(actual, Is.EquivalentTo(expected));
            }
        }

        [Test]
        [Category("Resolver")]
        [Category("Interfaces")]
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

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var declaration = declarations.Single(decl => decl.DeclarationType == DeclarationType.Variable && decl.IdentifierName.Equals("Bar"));

                var expected = declarations.Where(decl => decl.IdentifierName.Equals("TestInterface_Bar") &&
                                                          decl.DeclarationType == DeclarationType.PropertyLet ||
                                                          decl.DeclarationType == DeclarationType.PropertyGet).ToList();
                var actual = state.DeclarationFinder.FindInterfaceImplementationMembers(declaration).ToList();
                var results = actual.Count;

                Assert.AreEqual(expected.Count, results, "Expected {0} Declarations, received {1}", expected.Count, results);
                Assert.That(actual, Is.EquivalentTo(expected));
            }
        }

        [Test]
        [Category("Resolver")]
        [Category("Interfaces")]
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

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var declaration = declarations.Single(decl => decl.DeclarationType == DeclarationType.Variable && decl.IdentifierName.Equals("Bar"));

                var expected = declarations.Where(decl => decl.IdentifierName.Equals("TestInterface_Bar") &&
                                                          decl.DeclarationType == DeclarationType.PropertySet ||
                                                          decl.DeclarationType == DeclarationType.PropertyGet).ToList();
                var actual = state.DeclarationFinder.FindInterfaceImplementationMembers(declaration).ToList();
                var results = actual.Count;

                Assert.AreEqual(expected.Count, results, "Expected {0} Declarations, received {1}", expected.Count, results);
                Assert.That(actual, Is.EquivalentTo(expected));
            }
        }

        [Test]
        [Category("Resolver")]
        [Category("Interfaces")]
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

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var declaration = declarations.Single(decl => decl.DeclarationType == DeclarationType.PropertyGet && decl.IdentifierName.Equals("TestInterface_Foo"));

                var expected = declarations.Single(decl => decl.DeclarationType == DeclarationType.PropertyGet && decl.IdentifierName.Equals("Foo"));
                var actual = (declaration as ModuleBodyElementDeclaration)?.InterfaceMemberImplemented;

                Assert.AreEqual(expected, actual, "Expected {0}, resolved to {1}", expected, actual);
            }
        }

        [Test]
        [Category("Resolver")]
        [Category("Interfaces")]
        public void FindFindInterfaceMemberParameterNamesIgnored()
        {
            var intrface =
                @"Option Explicit

Public Property Get Foo(Bar As Long) As Long
End Property

Public Property Let Foo(Bar As Long, Baz As Long)
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

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var declaration = declarations.Single(decl => decl.DeclarationType == DeclarationType.PropertyLet && decl.IdentifierName.Equals("TestInterface_Foo"));

                var expected = declarations.Single(decl => decl.DeclarationType == DeclarationType.PropertyLet && decl.IdentifierName.Equals("Foo"));
                var actual = (declaration as ModuleBodyElementDeclaration)?.InterfaceMemberImplemented;

                Assert.AreEqual(expected, actual, "Expected {0}, resolved to {1}", expected, actual);
            }
        }

        [Test]
        [Category("Resolver")]
        [Category("Interfaces")]
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

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var declaration = declarations.Single(decl => decl.DeclarationType == DeclarationType.PropertyGet && decl.IdentifierName.Equals("TestInterface_Foo"));

                var actual = (declaration as ModuleBodyElementDeclaration)?.InterfaceMemberImplemented;

                Assert.IsNull(actual, "Expected null, resolved to {0}", actual);
            }
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

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var expected = declarations.Single(decl => decl.IdentifierName.Equals("Foo") && decl.DeclarationType == DeclarationType.PropertyGet);
                var implementing = declarations.Single(decl => decl.IdentifierName.Equals("TestInterface_Foo") && decl.DeclarationType == DeclarationType.PropertyGet);

                var actual = (implementing as ModuleBodyElementDeclaration)?.InterfaceMemberImplemented;

                Assert.AreEqual(expected, actual, "Expected {0}, resolved to {1}", expected, actual);
            }
        }

        [Test]
        [Category("Resolver")]
        [Category("Interfaces")]
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

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var expected = declarations.Single(decl => decl.IdentifierName.Equals("Foo") && decl.DeclarationType == DeclarationType.Variable);
                var implementing = declarations.Single(decl => decl.IdentifierName.Equals("TestInterface_Foo") && decl.DeclarationType == DeclarationType.PropertyGet);

                var actual = (implementing as ModuleBodyElementDeclaration)?.InterfaceMemberImplemented;

                Assert.AreEqual(expected, actual, "Expected {0}, resolved to {1}", expected, actual);
            }
        }

        [Test]
        [Category("Resolver")]
        [Category("Interfaces")]
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

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var expected = declarations.Single(decl => decl.IdentifierName.Equals("Foo") && decl.DeclarationType == DeclarationType.Variable);
                var implementing = declarations.Single(decl => decl.IdentifierName.Equals("TestInterface_Foo") && decl.DeclarationType == DeclarationType.PropertyLet);

                var actual = (implementing as ModuleBodyElementDeclaration)?.InterfaceMemberImplemented;

                Assert.AreEqual(expected, actual, "Expected {0}, resolved to {1}", expected, actual);
            }
        }

        [Test]
        [Category("Resolver")]
        [Category("Interfaces")]
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

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var implementing = declarations.Single(decl => decl.IdentifierName.Equals("TestInterface_Foo") && decl.DeclarationType == DeclarationType.PropertySet);

                Assert.IsFalse((implementing as ModuleBodyElementDeclaration)?.IsInterfaceImplementation);
            }
        }

        [Test]
        [Category("Resolver")]
        [Category("Interfaces")]
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

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var expected = declarations.Single(decl => decl.IdentifierName.Equals("Foo") && decl.DeclarationType == DeclarationType.Variable);
                var implementing = declarations.Single(decl => decl.IdentifierName.Equals("TestInterface_Foo") && decl.DeclarationType == DeclarationType.PropertySet);

                var actual = (implementing as ModuleBodyElementDeclaration)?.InterfaceMemberImplemented;

                Assert.AreEqual(expected, actual, "Expected {0}, resolved to {1}", expected, actual);
            }
        }

        [Test]
        [Category("Resolver")]
        [Category("Interfaces")]
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

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var implementing = declarations.Single(decl => decl.IdentifierName.Equals("TestInterface_Foo") && decl.DeclarationType == DeclarationType.PropertyLet);

                Assert.IsFalse((implementing as ModuleBodyElementDeclaration)?.IsInterfaceImplementation);
            }
        }

        [Test]
        [Category("Resolver")]
        [Category("Interfaces")]
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

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var expected = declarations.Single(decl => decl.IdentifierName.Equals("Foo") && decl.DeclarationType == DeclarationType.Variable);
                var setter = declarations.Single(decl => decl.IdentifierName.Equals("TestInterface_Foo") && decl.DeclarationType == DeclarationType.PropertySet);
                var letter = declarations.Single(decl => decl.IdentifierName.Equals("TestInterface_Foo") && decl.DeclarationType == DeclarationType.PropertyLet);

                var actualSetter = (setter as ModuleBodyElementDeclaration)?.InterfaceMemberImplemented;
                var actualLetter = (letter as ModuleBodyElementDeclaration)?.InterfaceMemberImplemented;

                Assert.AreEqual(expected, actualSetter, "Expected {0}, resolved to {1}", expected, actualSetter);
                Assert.AreEqual(expected, actualLetter, "Expected {0}, resolved to {1}", expected, actualLetter);
            }
        }

        private static ClassModuleDeclaration GetTestClassModule(Declaration projectDeclatation, string name, bool isExposed = false)
        {
            var qualifiedClassModuleMemberName = new QualifiedMemberName(StubQualifiedModuleName(name), name);
            var classModuleAttributes = new Attributes();
            if (isExposed)
            {
                classModuleAttributes.AddExposedClassAttribute();
            }
            return new ClassModuleDeclaration(qualifiedClassModuleMemberName, projectDeclatation, name, true, null, classModuleAttributes);
        }

        private static ProjectDeclaration GetTestProject(string name)
        {
            var qualifiedProjectName = new QualifiedMemberName(StubQualifiedModuleName("proj"), name);
            return new ProjectDeclaration(qualifiedProjectName, name, true);
        }

        private static QualifiedModuleName StubQualifiedModuleName(string name)
        {
            return new QualifiedModuleName("dummy", "dummy", name);
        }

        private static FunctionDeclaration GetTestFunction(Declaration moduleDeclatation, string name, Accessibility functionAccessibility)
        {
            var qualifiedFunctionMemberName = new QualifiedMemberName(moduleDeclatation.QualifiedName.QualifiedModuleName, name);
            return new FunctionDeclaration(qualifiedFunctionMemberName, moduleDeclatation, moduleDeclatation, "test", null, "test", functionAccessibility, null, null, Selection.Home, false, true, null, null);
        }

        private static void AddReference(Declaration toDeclaration, Declaration fromModuleDeclaration, ParserRuleContext context = null)
        {
            toDeclaration.AddReference(toDeclaration.QualifiedName.QualifiedModuleName, fromModuleDeclaration, fromModuleDeclaration, context, toDeclaration.IdentifierName, toDeclaration, Selection.Home, new List<Rubberduck.Parsing.Annotations.IParseTreeAnnotation>());
        }
    }
}