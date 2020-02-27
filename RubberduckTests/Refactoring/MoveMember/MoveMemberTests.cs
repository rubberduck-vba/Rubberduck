using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.Refactorings.MoveMember.Extensions;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using System.Collections.Generic;
using System.Linq;

namespace RubberduckTests.Refactoring.MoveMember
{
    [TestFixture]
    public class MoveMemberTests : MoveMemberTestsBase
    {
        [TestCase("Foo(ByVal arg1 As Single, ByVal arg2 As Single, arg3 As Single)", "ByRef arg3")]
        [TestCase("Foo(arg1 As Single, ByVal arg2 As Single, arg3 As Single)", "ByRef arg1")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ImprovedArgListTest(string input, string expectedToContain)
        {
            var member = "Foo";
            var source =
$@"
Option Explicit

Public Sub {input}
End Sub
";

            var vbeStub = MockVbeBuilder.BuildFromSingleModule(source, Rubberduck.VBEditor.SafeComWrappers.ComponentType.StandardModule, out _);
            var result = MoveMemberTestSupport.ParseAndTest(vbeStub.Object, ThisTest);

            string ThisTest(RubberduckParserState state)
            {
                var target = state.DeclarationFinder.MatchName(member).Single();
                return target is IParameterizedDeclaration memberWithParams
                                ? memberWithParams.BuildFullyDefinedArgumentList()
                                : "()";
            }

            StringAssert.Contains(expectedToContain, result);
        }

        [TestCase("Foo(ByVal arg1 As Single, ByVal arg2 As Single, arg3 As Single) As Single", "ByRef arg3")]
        [TestCase("Foo(arg1 As Single, ByVal arg2 As Single, arg3 As Single) As Single", "ByRef arg1")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ImprovedArgListTestPropertyGet(string input, string expectedToContain)
        {
            var member = "Foo";
            var source =
$@"
Option Explicit

Public Property Get {input}
End Property
";

            var vbeStub = MockVbeBuilder.BuildFromSingleModule(source, Rubberduck.VBEditor.SafeComWrappers.ComponentType.StandardModule, out _);
            var result = MoveMemberTestSupport.ParseAndTest(vbeStub.Object, ThisTest);

            string ThisTest(RubberduckParserState state)
            {
                var target = state.DeclarationFinder.MatchName(member).Single();
                return target is IParameterizedDeclaration memberWithParams
                                ? memberWithParams.BuildFullyDefinedArgumentList()
                                : "()";
            }

            StringAssert.Contains(expectedToContain, result);
        }

        [TestCase("Foo(ByVal arg1 As Single, ByVal arg2 As Single, arg3 As Single)", "ByVal arg3")]
        [TestCase("Foo(arg1 As Single, ByVal arg2 As Single, ByRef arg3 As Single)", "ByRef arg1", "ByVal arg3")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ImprovedArgListTestPropertyLet(string input, params string[] expectedToContain)
        {
            var member = "Foo";
            var source =
$@"
Option Explicit

Public Property Let {input}
End Property
";

            var vbeStub = MockVbeBuilder.BuildFromSingleModule(source, Rubberduck.VBEditor.SafeComWrappers.ComponentType.StandardModule, out _);
            var result = MoveMemberTestSupport.ParseAndTest(vbeStub.Object, ThisTest);

            string ThisTest(RubberduckParserState state)
            {
                var target = state.DeclarationFinder.MatchName(member).Single();
                return target is IParameterizedDeclaration memberWithParams
                                ? memberWithParams.BuildFullyDefinedArgumentList()
                                : "()";
            }

            foreach (var expected in expectedToContain)
            {
                StringAssert.Contains(expected, result);
            }
        }

        [TestCase("Foo(arg1, ByVal arg2 As Single, arg3 As Single)", "ByRef arg1 As Variant", "ByVal arg3")]
        [TestCase("Foo(arg1, ByVal arg2, arg3 As Single)", "ByRef arg1 As Variant", "ByVal arg2 As Variant", "ByVal arg3")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ImprovedArgListTestImplicitTypes(string input, params string[] expectedToContain)
        {
            var member = "Foo";
            var source =
$@"
Option Explicit

Public Property Let {input}
End Property
";

            var vbeStub = MockVbeBuilder.BuildFromSingleModule(source, ComponentType.StandardModule, out _);
            var result = MoveMemberTestSupport.ParseAndTest(vbeStub.Object, ThisTest);

            string ThisTest(RubberduckParserState state)
            {
                var target = state.DeclarationFinder.MatchName(member).Single();
                return target is IParameterizedDeclaration memberWithParams
                                ? memberWithParams.BuildFullyDefinedArgumentList()
                                : "()";
            }

            foreach (var expected in expectedToContain)
            {
                StringAssert.Contains(expected, result);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MoveCandidatesPropertiesGroupedTogether()
        {
            var member = "Foo";
            var source =
$@"
Option Explicit

Public Property Let Foo(arg As Long)
End Property

Public Property Get Foo() As Long
End Property

Public Property Let Bar(arg As Long)
End Property

Public Property Get Bar() As Long
End Property

Public Property Get Baz() As Long
End Property
";

            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, (member, DeclarationType.PropertyGet), sourceContent: source);


            var vbeStub = MoveMemberTestSupport.BuildVBEStub(moveDefinition, source);
            var moveCandidates = MoveMemberTestSupport.ParseAndTest(vbeStub, ThisTest);


            IEnumerable<IMoveableMemberSet> ThisTest(RubberduckParserState state, IVBE vbe, IRewritingManager rewritingManager)
            {
                var model = MoveMemberTestSupport.CreateModelAndDefineMove(vbe, moveDefinition, state, rewritingManager);
                return model.MoveableMembers;
            }

            Assert.AreEqual(3, moveCandidates.Count());
        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MoveCandidatesPreview()
        {
            var memberToMove = ("Foo", DeclarationType.PropertyGet);
            var source =
$@"
Option Explicit

Public Property Let Foo(arg As Long)
End Property

Public Property Get Foo() As Long
End Property

Public Property Let Bar(arg As Long)
End Property

Public Property Get Bar() As Long
End Property

Public Property Get Baz() As Long
End Property
";

            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, memberToMove, sourceContent: source, createNewModule: true);
            var preview = RetrievePreviewAfterUserInput(moveDefinition, source, memberToMove);

            Assert.IsTrue(MoveMemberTestSupport.OccursOnce("Let Foo(", preview), "String occurs more than once");
            Assert.IsTrue(MoveMemberTestSupport.OccursOnce("Get Foo(", preview), "String occurs more than once");
        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MultiplePropertyGroupsReferenceSameVariable()
        {
            var member = "Foo";
            var source =
$@"
Option Explicit

Private Const mFoo As Long = 10

Public Property Get Foo() As Long
    Foo = mFoo
End Property

Public Property Get FooTimes2() As Long 
    FooTimes2 = mFoo * 2
End Property
";

            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, (member, DeclarationType.PropertyGet), sourceContent: source);


            var refactoredCode = RefactoredCode(moveDefinition, source, null, null, false, ("FooTimes2", DeclarationType.PropertyGet));

            StringAssert.Contains("Get Foo()", refactoredCode.Destination);
            StringAssert.Contains("Get FooTimes2()", refactoredCode.Destination);
            StringAssert.Contains("Private Const mFoo As Long = 10", refactoredCode.Destination);
        }

        [TestCase("foo", "foo1")]
        [TestCase("foo1", "foo2")]
        [TestCase("foo123", "foo124")]
        [TestCase("f67oo3", "f67oo4")]
        [TestCase("foo0", "foo1")]
        [TestCase("", "1")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void IdentifierNameIncrementing(string input, string expected)
        {
            var actual = input.IncrementIdentifier();
            Assert.AreEqual(expected, actual);
        }
    }
}