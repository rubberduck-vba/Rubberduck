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
    public class ParameterDeclarationTests
    {
        [Test]
        [Category("Resolver")]
        public void ParametersHaveDeclarationTypeParameter()
        {
            var paramter = GetTestParameter("testParam", false, false, false);

            Assert.IsTrue(paramter.DeclarationType.HasFlag(DeclarationType.Parameter));
        }

            private static ParameterDeclaration GetTestParameter(string name, bool isOptional, bool isByRef, bool isParamArray)
            {
                var qualifiedParameterName = new QualifiedMemberName(StubQualifiedModuleName(), name);
                return new ParameterDeclaration(qualifiedParameterName, null, "test", null,"test", isOptional,isByRef, false, isParamArray);
            }

                private static QualifiedModuleName StubQualifiedModuleName()
                {
                    return new QualifiedModuleName("dummy", "dummy", "dummy");
                }


        [Test]
        [Category("Resolver")]
        public void ParametersHaveImpliciteAccessibility()
        {
            var paramter = GetTestParameter("testParam", false, false, false);

            Assert.IsTrue(paramter.Accessibility.HasFlag(Accessibility.Implicit));
        }


        [Test]
        [Category("Resolver")]
        public void IsParamArrayCanBeSetPublicly()
        {
            var paramter = GetTestParameter("testParam", false, false, false);
            paramter.IsParamArray = true;

            Assert.IsTrue(paramter.IsParamArray);
        }

        [Test]
        [Category("Resolver")]
        public void DefaultParameterValueIsResolved()
        {
            var code =
@"Option Explicit

Public Sub Foo(Optional bar As Long = 42)
    Debug.Print bar
End Sub
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("ParameterTest", ProjectProtection.Unprotected)
                .AddComponent("UnderTest", ComponentType.StandardModule, code, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            var parser = MockParser.Create(vbe.Object);
            parser.Parse(new CancellationTokenSource());

            var member = (ParameterDeclaration)parser.State.DeclarationFinder.UserDeclarations(DeclarationType.Parameter).Single(x => x.IdentifierName.Equals("bar"));
            var expected = "42";
            var actual = member.DefaultValue;

            Assert.AreEqual(expected, actual, "Expected {0}, resolved to {1}", expected, actual);
        }

        [Test]
        [Category("Resolver")]
        public void DefaultParameterValueIsResolvedForExpression()
        {
            var code =
@"Option Explicit

Public Sub Foo(Optional bar As Long = 6 * 7)
    Debug.Print bar
End Sub
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("ParameterTest", ProjectProtection.Unprotected)
                .AddComponent("UnderTest", ComponentType.StandardModule, code, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            var parser = MockParser.Create(vbe.Object);
            parser.Parse(new CancellationTokenSource());

            var member = (ParameterDeclaration)parser.State.DeclarationFinder.UserDeclarations(DeclarationType.Parameter).Single(x => x.IdentifierName.Equals("bar"));
            var expected = "6 * 7";
            var actual = member.DefaultValue;

            Assert.AreEqual(expected, actual, "Expected {0}, resolved to {1}", expected, actual);
        }

        [Test]
        [Category("Resolver")]
        public void NoDefaultOptionalParameterDefaultValueIsEmptyString()
        {
            var code =
@"Option Explicit

Public Sub Foo(Optional bar As Long)
    Debug.Print bar
End Sub
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("ParameterTest", ProjectProtection.Unprotected)
                .AddComponent("UnderTest", ComponentType.StandardModule, code, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            var parser = MockParser.Create(vbe.Object);
            parser.Parse(new CancellationTokenSource());

            var member = (ParameterDeclaration)parser.State.DeclarationFinder.UserDeclarations(DeclarationType.Parameter).Single(x => x.IdentifierName.Equals("bar"));
            var expected = string.Empty;
            var actual = member.DefaultValue;

            Assert.AreEqual(expected, actual, "Expected string.Empty, resolved to {0}", actual);
        }

        [Test]
        [Category("Resolver")]
        public void RequiredParameterDefaultValueIsEmptyString()
        {
            var code =
@"Option Explicit

Public Sub Foo(bar As Long)
    Debug.Print bar
End Sub
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("ParameterTest", ProjectProtection.Unprotected)
                .AddComponent("UnderTest", ComponentType.StandardModule, code, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            var parser = MockParser.Create(vbe.Object);
            parser.Parse(new CancellationTokenSource());

            var member = (ParameterDeclaration)parser.State.DeclarationFinder.UserDeclarations(DeclarationType.Parameter).Single(x => x.IdentifierName.Equals("bar"));
            var expected = string.Empty;
            var actual = member.DefaultValue;

            Assert.AreEqual(expected, actual, "Expected string.Empty, resolved to {0}", actual);
        }

        [Test]
        [Category("Resolver")]
        public void ImplicitTypedParameterDefaultValueIsEmptyString()
        {
            var code =
@"Option Explicit

Public Sub Foo(bar)
    Debug.Print bar
End Sub
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("ParameterTest", ProjectProtection.Unprotected)
                .AddComponent("UnderTest", ComponentType.StandardModule, code, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            var parser = MockParser.Create(vbe.Object);
            parser.Parse(new CancellationTokenSource());

            var member = (ParameterDeclaration)parser.State.DeclarationFinder.UserDeclarations(DeclarationType.Parameter).Single(x => x.IdentifierName.Equals("bar"));
            var expected = string.Empty;
            var actual = member.DefaultValue;

            Assert.AreEqual(expected, actual, "Expected string.Empty, resolved to {0}", actual);
        }
    }
}
