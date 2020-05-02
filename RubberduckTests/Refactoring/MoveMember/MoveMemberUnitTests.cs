using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.MoveMember;
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
                var resolver = new MoveMemberTestsResolver(state, rewritingManager);
                var factory = resolver.Resolve<IMoveableMemberSetsFactory>();

                var result =  factory.Create(targets.First()).ToDictionary(key => key.IdentifierName).Values;
                Assert.AreEqual(3, result.Count());
            }
        }

        //[TestCase("fizz", "fizz1")]
        //[TestCase("fizz1", "fizz2")]
        //[TestCase("fizz123", "fizz124")]
        //[TestCase("f67oo3", "f67oo4")]
        //[TestCase("fizz0", "fizz1")]
        //[TestCase("", "1")]
        //[Category("Refactorings")]
        //[Category("MoveMember")]
        //public void IdentifierNameIncrementing(string input, string expected)
        //{
        //    var actual = input.IncrementIdentifier();
        //    Assert.AreEqual(expected, actual);
        //}

        [TestCase("mTest")]
        [TestCase("mTest", "mTest1")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MoveMemberModelFactoryCreateOverloadsNamedDestination(params string[] fieldsToMove)
        {
            var endpoints = MoveEndpoints.StdToStd;
            var source =
$@"
Option Explicit

Public mTest As Long

Public mTest1 As Long
";
            (RubberduckParserState state, IRewritingManager rewritingManager) = CreateAndParse(MoveEndpoints.StdToStd, source, string.Empty);
            using (state)
            {
                var useDestinationName = true;
                var destinationContent = ExecuteMoveMemberRefactoringAction(state, rewritingManager, endpoints, fieldsToMove, useDestinationName);

                StringAssert.Contains("Public mTest As Long", destinationContent);
                if (fieldsToMove.Count() > 1)
                {
                    StringAssert.Contains("Public mTest1 As Long", destinationContent);
                }
            }
        }

        [TestCase("mTest")]
        [TestCase("mTest", "mTest1")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MoveMemberModelFactoryCreateOverloadsDestinationDeclaration(params string[] fieldsToMove)
        {
            var endpoints = MoveEndpoints.StdToStd;
            var source =
$@"
Option Explicit

Public mTest As Long

Public mTest1 As Long
";
            (RubberduckParserState state, IRewritingManager rewritingManager) = CreateAndParse(MoveEndpoints.StdToStd, source, string.Empty);
            using (state)
            {
                var useDestinationName = false;
                var destinationContent = ExecuteMoveMemberRefactoringAction(state, rewritingManager, endpoints, fieldsToMove, useDestinationName);

                StringAssert.Contains("Public mTest As Long", destinationContent);
                if (fieldsToMove.Count() > 1)
                {
                    StringAssert.Contains("Public mTest1 As Long", destinationContent);
                }
            }
        }

        [TestCase("mTest")]
        [TestCase("mTest", "mTest1")]
        [TestCase("mTest1", "mTest2")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MoveMemberModelFactoryCreateOverloadsUsesDestinationNameThrows(params string[] fieldsToMove)
        {
            var endpoints = MoveEndpoints.StdToStd;
            var source =
$@"
Option Explicit

Private Type TestType
    FirstVal As Long
End Type

Public mTest As TestType

Private mINeedMyType As TestType

Public mTest1 As Long

Private mTest2 As Long
";
            (RubberduckParserState state, IRewritingManager rewritingManager) = CreateAndParse(MoveEndpoints.StdToStd, source, string.Empty);
            using (state)
            {
                var useDestinationName = false;
                Assert.Throws<MoveMemberUnsupportedMoveException>(() => ExecuteMoveMemberRefactoringAction(state, rewritingManager, endpoints, fieldsToMove, useDestinationName));
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MoveMemberModelFactoryNullTargetThrows()
        {
            var source =
$@"
Option Explicit

Public mTest As Long
";
            (RubberduckParserState state, IRewritingManager rewritingManager) = CreateAndParse(MoveEndpoints.StdToStd, source, string.Empty);
            using (state)
            {
                var resolver = new MoveMemberTestsResolver(state, rewritingManager);
                var modelFactory = resolver.Resolve<IMoveMemberModelFactory>();

                Declaration target = null;
                Assert.Throws<TargetDeclarationIsNullException>(() => modelFactory.Create(target));
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MoveMemberModelFactoryInvalidTargetTypeThrows()
        {
            var source =
$@"
Option Explicit

Public Enum ETest
    FirstValue = 0
    SecondValue
End Enum
";
            (RubberduckParserState state, IRewritingManager rewritingManager) = CreateAndParse(MoveEndpoints.StdToStd, source, string.Empty);
            using (state)
            {
                var target = state.DeclarationFinder.DeclarationsWithType(DeclarationType.EnumerationMember)
                    .Where(d => d.IdentifierName == "SecondValue");

                var resolver = new MoveMemberTestsResolver(state, rewritingManager);
                var modelFactory = resolver.Resolve<IMoveMemberModelFactory>();

                Assert.Throws<MoveMemberUnsupportedMoveException>(() => modelFactory.Create(target));
            }
        }


        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MoveMemberToClassModuleThrow()
        {
            var endpoints = MoveEndpoints.StdToClass;
            var source =
$@"
Option Explicit

Public Sub NoCanDo()
End Sub
";
            (RubberduckParserState state, IRewritingManager rewritingManager) = CreateAndParse(MoveEndpoints.StdToClass, source, string.Empty);
            using (state)
            {
                var target = state.DeclarationFinder.DeclarationsWithType(DeclarationType.Procedure)
                    .Where(d => d.IdentifierName == "NoCanDo");

                var resolver = new MoveMemberTestsResolver(state, rewritingManager);
                var modelFactory = resolver.Resolve<IMoveMemberModelFactory>();

                var model = modelFactory.Create(target);
                model.ChangeDestination(endpoints.DestinationModuleName(), endpoints.DestinationComponentType());

                var refactoringAction = resolver.Resolve<MoveMemberToExistingModuleRefactoringAction>();

                Assert.Throws<MoveMemberUnsupportedMoveException>(() => refactoringAction.Refactor(model, rewritingManager.CheckOutCodePaneSession()));
            }
        }

        private string ExecuteMoveMemberRefactoringAction(RubberduckParserState state, IRewritingManager rewritingManager, MoveEndpoints endpoints, string[] fieldsToMove, bool useDestinationName)
        {
            var resolver = new MoveMemberTestsResolver(state, rewritingManager);
            var targets = state.DeclarationFinder.DeclarationsWithType(DeclarationType.Variable)
                .Where(d => fieldsToMove.Contains(d.IdentifierName));

            var destination = state.DeclarationFinder.UserDeclarations(DeclarationType.ProceduralModule)
                .Where(d => d.IdentifierName.Equals(endpoints.DestinationModuleName())).OfType<ModuleDeclaration>().Single();

            var modelFactory = resolver.Resolve<IMoveMemberModelFactory>();
            var refactoringAction = resolver.Resolve<MoveMemberToExistingModuleRefactoringAction>();
            var rewriteSession = rewritingManager.CheckOutCodePaneSession();

            MoveMemberModel model;
            if (fieldsToMove.Count() == 1)
            {
                if (useDestinationName)
                {
                    model = modelFactory.Create(targets.First(), endpoints.DestinationModuleName());
                }
                else
                {
                    model = modelFactory.Create(targets.First(), destination);
                }
            }
            else
            {
                if (useDestinationName)
                {
                    model = modelFactory.Create(targets, endpoints.DestinationModuleName());
                }
                else
                {
                    model = modelFactory.Create(targets, destination);
                }
            }

            refactoringAction.Refactor(model);
            return rewriteSession.CheckOutModuleRewriter(destination.QualifiedModuleName)
                                        .GetText();
        }

        private (RubberduckParserState, IRewritingManager) CreateAndParse(MoveEndpoints endpoints, string sourceContent, string destinationContent)
        {
            var modules = endpoints.ToModulesTuples(sourceContent, destinationContent);
            var vbe = MockVbeBuilder.BuildFromModules(modules).Object;
            return MockParser.CreateAndParseWithRewritingManager(vbe);
        }

        public static string TestForImprovedArgumentList(string sourceContent, string targetID)
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