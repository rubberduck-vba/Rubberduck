using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
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
    public class MoveMemberUnitTests
    {
        //No Changes unless there is a UDT passed ByVal
        [TestCase("Fizz(ByVal arg1 As Single, ByVal arg2 As Single, arg3 As Single)", "ByVal arg1 As Single, ByVal arg2 As Single, arg3 As Single")]
        [TestCase("Fizz(arg1 As Single, ByVal arg2 As Single, arg3 As Single)", "arg1 As Single, ByVal arg2 As Single, arg3 As Single")]
        [TestCase("Fizz(ByVal arg1 As TestType, ByVal arg2 As Single, arg3 As Single)", "ByRef arg1 As TestType, ByVal arg2 As Single, arg3 As Single")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ImprovedArgListSub(string input, string expectedToContain)
        {
            var member = "Fizz";
            var source =
$@"
Option Explicit

Public Type TestType
    MemberValue As Long
End Type

Public Sub {input}
End Sub
";

            var vbeStub = MockVbeBuilder.BuildFromSingleModule(source, ComponentType.StandardModule, out _);
            var result = MoveMemberTestSupport.ParseAndTest(vbeStub.Object, ThisTest);

            string ThisTest(RubberduckParserState state)
            {
                var target = state.DeclarationFinder.MatchName(member).Single();
                return target is ModuleBodyElementDeclaration mbed
                                ? mbed.ImprovedArgumentList()
                                : string.Empty;
            }

            StringAssert.Contains(expectedToContain, result);
        }

        [TestCase("Fizz(ByVal arg1 As Single, ByVal arg2 As Single, arg3 As Single) As Single", "ByVal arg1 As Single, ByVal arg2 As Single, arg3 As Single")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ImprovedArgListPropertyGet(string input, string expectedToContain)
        {
            var member = "Fizz";
            var source =
$@"
Option Explicit

Public Property Get {input}
End Property
";

            var vbeStub = MockVbeBuilder.BuildFromSingleModule(source, ComponentType.StandardModule, out _);
            var result = MoveMemberTestSupport.ParseAndTest(vbeStub.Object, ThisTest);

            string ThisTest(RubberduckParserState state)
            {
                var target = state.DeclarationFinder.MatchName(member).Single();
                return target is ModuleBodyElementDeclaration mbed
                                ? mbed.ImprovedArgumentList()
                                : string.Empty;
            }

            StringAssert.Contains(expectedToContain, result);
        }

        [TestCase("Fizz(ByVal arg1 As Single, ByVal arg2 As Single, arg3 As Single)", "ByVal arg1 As Single, ByVal arg2 As Single, ByVal arg3 As Single")]
        [TestCase("Fizz(arg1 As Single, ByVal arg2 As Single, ByRef arg3 As Single)", "arg1 As Single, ByVal arg2 As Single, ByVal arg3 As Single")]
        [TestCase("Fizz(arg1 As Single, ByVal arg2 As TestType, ByRef arg3 As Single)", "arg1 As Single, ByRef arg2 As TestType, ByVal arg3 As Single")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ImprovedArgListPropertyLet(string input, params string[] expectedToContain)
        {
            var member = "Fizz";
            var source =
$@"
Option Explicit

Public Type TestType
    MemberValue As Long
End Type

Public Property Let {input}
End Property
";

            var vbeStub = MockVbeBuilder.BuildFromSingleModule(source, ComponentType.StandardModule, out _);
            var result = MoveMemberTestSupport.ParseAndTest(vbeStub.Object, ThisTest);

            string ThisTest(RubberduckParserState state)
            {
                var target = state.DeclarationFinder.MatchName(member).Single();
                return target is ModuleBodyElementDeclaration mbed
                                ? mbed.ImprovedArgumentList()
                                : string.Empty;
            }

            foreach (var expected in expectedToContain)
            {
                StringAssert.Contains(expected, result);
            }
        }

        [TestCase("Fizz(arg1, ByVal arg2 As Single, arg3 As Single)", "arg1 As Variant", "ByVal arg3")]
        [TestCase("Fizz(arg1, ByVal arg2, arg3 As Single)", "arg1 As Variant", "ByVal arg2 As Variant", "ByVal arg3")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ImprovedArgListImplicitTypes(string input, params string[] expectedToContain)
        {
            var member = "Fizz";
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
                return target is ModuleBodyElementDeclaration mbed
                                ? mbed.ImprovedArgumentList()
                                : string.Empty;
            }

            foreach (var expected in expectedToContain)
            {
                StringAssert.Contains(expected, result);
            }
        }

        [TestCase("Fizz", "arg1 As Long, Optional arg2 As Long = 4")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ImprovedArgListOptionalParam(string member, string expectedDisplay)
        {
            var source =
$@"
Option Explicit

Public Sub Fizz(arg1 As Long, Optional arg2 As Long = 4)
End Sub
";

            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, (member, DeclarationType.PropertyGet), sourceContent: source);

            var vbeStub = MoveMemberTestSupport.BuildVBEStub(moveDefinition, source);
            var displaySignature = MoveMemberTestSupport.ParseAndTest(vbeStub, ThisTest);

            StringAssert.AreEqualIgnoringCase(expectedDisplay, displaySignature);


            string ThisTest(RubberduckParserState state, IVBE vbe, IRewritingManager rewritingManager)
            {
                var target = state.DeclarationFinder.MatchName(member).Single();
                return target is ModuleBodyElementDeclaration mbed
                                ? mbed.ImprovedArgumentList()
                                : string.Empty;
            }
        }

        [TestCase("Fizz", "arg1 As Long, ParamArray moreArgs() As Variant")]
        [TestCase("Bizz", "arg1 As Long, ParamArray moreArgs() As Variant")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ImprovedArgListParamArray(string member, string expectedDisplay)
        {
            var source =
$@"
Option Explicit

Public Sub Fizz(arg1 As Long, ParamArray moreArgs() As Variant)
End Sub

Public Sub Bizz(arg1 As Long, ParamArray moreArgs())
End Sub
";

            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, (member, DeclarationType.PropertyGet), sourceContent: source);

            var vbeStub = MoveMemberTestSupport.BuildVBEStub(moveDefinition, source);
            var displaySignature = MoveMemberTestSupport.ParseAndTest(vbeStub, ThisTest);

            StringAssert.AreEqualIgnoringCase(expectedDisplay, displaySignature);


            string ThisTest(RubberduckParserState state, IVBE vbe, IRewritingManager rewritingManager)
            {
                var target = state.DeclarationFinder.MatchName(member).Single();
                return target is ModuleBodyElementDeclaration mbed
                                ? mbed.ImprovedArgumentList()
                                : string.Empty;
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MoveCandidatesPropertiesGroupedTogether()
        {
            var member = "Fizz";
            var source =
$@"
Option Explicit

Public Property Let Fizz(arg As Long)
End Property

Public Property Get Fizz() As Long
End Property

Public Property Let Bizz(arg As Long)
End Property

Public Property Get Bizz() As Long
End Property

Public Property Get Bazz() As Long
End Property
";

            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, (member, DeclarationType.PropertyGet), sourceContent: source);


            var vbeStub = MoveMemberTestSupport.BuildVBEStub(moveDefinition, source);
            var moveCandidates = MoveMemberTestSupport.ParseAndTest(vbeStub, ThisTest);

            IEnumerable<IMoveableMemberSet> ThisTest(RubberduckParserState state, IVBE vbe, IRewritingManager rewritingManager)
            {
                var target = state.DeclarationFinder.MatchName(member);

                var model = MoveMemberTestSupport.CreateRefactoringModel(target.First(), state, rewritingManager);
                return model.MoveableMembers;
            }

            Assert.AreEqual(3, moveCandidates.Count());
        }

        [TestCase("fizz", "fizz1")]
        [TestCase("fizz1", "fizz2")]
        [TestCase("fizz123", "fizz124")]
        [TestCase("f67oo3", "f67oo4")]
        [TestCase("fizz0", "fizz1")]
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