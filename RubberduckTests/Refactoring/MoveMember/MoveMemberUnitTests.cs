using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.MoveMember.Extensions;
using RubberduckTests.Mocks;
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

            var result = TestForImprovedArgumentList(source, member);
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

            var result = TestForImprovedArgumentList(source, member);
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

            var result = TestForImprovedArgumentList(source, member);
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

            var result = TestForImprovedArgumentList(source, member);
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

            var displaySignature = TestForImprovedArgumentList(source, member);
            StringAssert.AreEqualIgnoringCase(expectedDisplay, displaySignature);
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

            var displaySignature = TestForImprovedArgumentList(source, member);
            StringAssert.AreEqualIgnoringCase(expectedDisplay, displaySignature);
        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void GroupsMoveCandidatesPropertiesTogether()
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

            var vbeStub = MockVbeBuilder.BuildFromSingleStandardModule(source, out _);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbeStub.Object);
            using (state)
            {
                var targets = state.DeclarationFinder.MatchName(member);
                var serviceLocator = new MoveMemberTestsResolver(state, rewritingManager);
                var factory = serviceLocator.Resolve<IMoveableMemberSetsFactory>();

                var result =  factory.Create(targets.First()).ToDictionary(key => key.IdentifierName).Values;
                Assert.AreEqual(3, result.Count());
            }
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

        public static string TestForImprovedArgumentList(string targetID, string sourceContent)
        {
            var vbeStub = MockVbeBuilder.BuildFromSingleStandardModule(sourceContent, out _).Object;
            var state = MockParser.CreateAndParse(vbeStub);
            using (state)
            {
                var target = state.DeclarationFinder.MatchName(targetID).Single();
                return target is ModuleBodyElementDeclaration mbed
                                ? mbed.ImprovedArgumentList()
                                : string.Empty;
            }
        }
    }
}