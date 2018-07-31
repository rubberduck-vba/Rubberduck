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
    public class ValuedDeclarationTests
    {
        [Test]
        [Category("Resolver")]
        public void ExpressionReturnsTheContructorInjectedValue()
        {
            var value = "testtest";
            var constantName =  new QualifiedMemberName(StubQualifiedModuleName(),"testConstant");
            var constantDeclaration = new ValuedDeclaration(constantName, null, "test", "test", null, "test", null, Accessibility.Implicit, DeclarationType.Constant, value, null, Selection.Home, true);

            Assert.AreEqual(value, constantDeclaration.Expression);
        }
            
            private static QualifiedModuleName StubQualifiedModuleName()
            {
                return new QualifiedModuleName("dummy", "dummy", "dummy");
            }

        [Test]
        [Category("Resolver")]
        public void ConstantValueIsResolvedWhenLiteral()
        {
            var code =
@"Option Explicit

Private Const FORTY_TWO = 42

Public Function FortyTwo() As Long
    FortyTwo = FORTY_TWO
End Function
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("ConstTest", ProjectProtection.Unprotected)
                .AddComponent("UnderTest", ComponentType.StandardModule, code, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            var parser = MockParser.Create(vbe.Object);
            parser.Parse(new CancellationTokenSource());

            var constant = (ValuedDeclaration)parser.State.DeclarationFinder.DeclarationsWithType(DeclarationType.Constant).Single();
            var expected = "42";
            var actual = constant.Expression;

            Assert.AreEqual(expected, actual, "Expected {0}, resolved to {1}", expected, actual);
        }

        [Test]
        [Category("Resolver")]
        public void ConstantValueIsResolvedWhenExpression()
        {
            var code =
@"Option Explicit

Private Const FORTY_TWO = 6 * 7

Public Function FortyTwo() As Long
    FortyTwo = FORTY_TWO
End Function
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("ConstTest", ProjectProtection.Unprotected)
                .AddComponent("UnderTest", ComponentType.StandardModule, code, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            var parser = MockParser.Create(vbe.Object);
            parser.Parse(new CancellationTokenSource());

            var constant = (ValuedDeclaration)parser.State.DeclarationFinder.DeclarationsWithType(DeclarationType.Constant).Single();
            var expected = "6 * 7";
            var actual = constant.Expression;

            Assert.AreEqual(expected, actual, "Expected {0}, resolved to {1}", expected, actual);
        }

        [Test]
        [Category("Resolver")]
        public void ConstantValueIsResolvedWhenOtherConstants()
        {
            var code =
@"Option Explicit

Private Const SEVEN = 7
Private Const SIX = 6
Private Const FORTY_TWO = SIX * SEVEN

Public Function FortyTwo() As Long
    FortyTwo = FORTY_TWO
End Function
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("ConstTest", ProjectProtection.Unprotected)
                .AddComponent("UnderTest", ComponentType.StandardModule, code, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            var parser = MockParser.Create(vbe.Object);
            parser.Parse(new CancellationTokenSource());

            var constant = (ValuedDeclaration)parser.State.DeclarationFinder.UserDeclarations(DeclarationType.Constant).Single(x => x.IdentifierName.Equals("FORTY_TWO"));
            var expected = "SIX * SEVEN";
            var actual = constant.Expression;

            Assert.AreEqual(expected, actual, "Expected {0}, resolved to {1}", expected, actual);
        }

        [Test]
        [Category("Resolver")]
        public void EnumValueIsResolved()
        {
            var code =
@"Option Explicit

Private Enum Foo
    Bar = 24
    Baz = 42
End Enum

Public Function FortyTwo() As Long
    Debug.Print Foo.Bar
End Function
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("ConstTest", ProjectProtection.Unprotected)
                .AddComponent("UnderTest", ComponentType.StandardModule, code, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            var parser = MockParser.Create(vbe.Object);
            parser.Parse(new CancellationTokenSource());

            var member = (ValuedDeclaration)parser.State.DeclarationFinder.UserDeclarations(DeclarationType.EnumerationMember).Single(x => x.IdentifierName.Equals("Bar"));
            var expected = "24";
            var actual = member.Expression;

            Assert.AreEqual(expected, actual, "Expected {0}, resolved to {1}", expected, actual);
        }

        [Test]
        [Category("Resolver")]
        public void EnumWithoutValueIsEmptyString()
        {
            var code =
@"Option Explicit

Private Enum Foo
    Bar
    Baz
End Enum

Public Function FortyTwo() As Long
    Debug.Print Foo.Bar
End Function
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("ConstTest", ProjectProtection.Unprotected)
                .AddComponent("UnderTest", ComponentType.StandardModule, code, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            var parser = MockParser.Create(vbe.Object);
            parser.Parse(new CancellationTokenSource());

            var member = (ValuedDeclaration)parser.State.DeclarationFinder.UserDeclarations(DeclarationType.EnumerationMember).Single(x => x.IdentifierName.Equals("Bar"));
            var expected = string.Empty;
            var actual = member.Expression;

            Assert.AreEqual(expected, actual, "Expected string.Empty, resolved to {0}", actual);
        }
    }
}
