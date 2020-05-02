using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.VBEditor.SafeComWrappers;
using System.Linq;

namespace RubberduckTests.Refactoring.MoveMember
{
    [TestFixture]
    public class MoveUDTToStdModuleTests : MoveMemberRefactoringActionTestSupportBase
    {
        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void PublicUDTMoveField_HasReferences()
        {
            var memberToMove = ("mFooBar", DeclarationType.Variable);
            var endpoints = MoveEndpoints.StdToStd;
            var source =
$@"
Option Explicit

Public Type MyTestType
    Foo As Long
    Bar As String
End Type

Public mFooBar As MyTestType

Public Function FooMath(arg1 As Long) As Long
    FooMath = arg1 * mFooBar.Foo
End Function

Public Function ConcatBar(arg1 As String) As String
    ConcatBar = arg1 & mFooBar.Bar
End Function";

            var callSiteModuleName = "CallSiteModule";

            var otherModuleReference =
    $@"
Option Explicit

Public Function MemberAccessFoo(arg1 As Long) As Long
    MemberAccessFoo = arg1 + {endpoints.SourceModuleName()}.mFooBar.Foo * 2
End Function

Public Function WithMemberAccessFoo(arg1 As Long) As Long
    Dim result As Long
    With {endpoints.SourceModuleName()}
        result = (.mFooBar.Foo + arg1) * 2
    End With
    WithMemberAccessFoo = result
End Function

Public Function WithMemberAccessFoo2(arg1 As Long) As Long
    Dim result As Long
    With {endpoints.SourceModuleName()}.mFooBar
        result2 = (.Foo + arg1) * 2
    End With
    WithMemberAccessFoo2 = result2
End Function

Public Function NonQualifiedFoo(arg1 As Long) As Long
    NonQualifiedFoo = (mFooBar.Foo + arg1) * 2
End Function
";
            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, 
                endpoints.ToSourceTuple(source),
                endpoints.ToDestinationTuple(string.Empty),
                (callSiteModuleName, otherModuleReference, ComponentType.StandardModule));

            var destinationDeclaration = "Public mFooBar As MyTestType";

            StringAssert.DoesNotContain("mFooBar As MyTestType", refactoredCode.Source);
            StringAssert.Contains(destinationDeclaration, refactoredCode.Destination);
            StringAssert.Contains($"{endpoints.DestinationModuleName()}.mFooBar.Foo", refactoredCode.Source);
            StringAssert.Contains($"{endpoints.DestinationModuleName()}.mFooBar.Bar", refactoredCode.Source);
            StringAssert.Contains($"{endpoints.DestinationModuleName()}.mFooBar", refactoredCode[callSiteModuleName]);
            StringAssert.Contains($"result = ({endpoints.DestinationModuleName()}.mFooBar.Foo + arg1) * 2", refactoredCode[callSiteModuleName]);
            StringAssert.Contains($"With {endpoints.DestinationModuleName()}.mFooBar", refactoredCode[callSiteModuleName]);
            StringAssert.Contains($"NonQualifiedFoo = ({endpoints.DestinationModuleName()}.mFooBar.Foo + arg1) * 2", refactoredCode[callSiteModuleName]);
        }

        [TestCase(MoveEndpoints.StdToStd, true)] //OK for Private UDTType exposed as Public Field
        [TestCase(MoveEndpoints.StdToClass, false)] //Not compilable for Private UDTType exposed as Public Field MS-VBAL 5.2.3.1
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void PrivateUDTMovePublicFieldFromStdModule(MoveEndpoints endpoints, bool expected)
        {
            (string targetID, DeclarationType DecType) = ("mFooBar", DeclarationType.Variable);
            var source =
$@"
Option Explicit

Private Type MyTestType
    Foo As Long
    Bar As String
End Type

Public mFooBar As MyTestType

Public Function FooMath(arg1 As Long) As Long
    FooMath = arg1 * mFooBar.Foo
End Function

Public Function ConcatBar(arg1 As String) As String
    ConcatBar = arg1 & mFooBar.Bar
End Function";

            var state = CreateAndParse(endpoints, source, string.Empty);
            using (state)
            {
                var resolver = new MoveMemberTestsResolver(state);
                var strategyFactory = resolver.Resolve<IMoveMemberStrategyFactory>();
                var target = state.DeclarationFinder.DeclarationsWithType(DecType).Where(d => d.IdentifierName == targetID).Single();
                var destination = state.DeclarationFinder.DeclarationsWithType(endpoints.ToDeclarationType()).Where(d => d.IdentifierName == endpoints.DestinationModuleName()).Single();
                var model = resolver.Resolve<IMoveMemberModelFactory>().Create(target, destination as ModuleDeclaration);

                Assert.AreEqual(expected, model.TryGetStrategy(out _));
            }
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void PublicUDTFieldReferencedByMovedFunction(string udtAccessibility)
        {
            var memberToMove = ("FizzMath", DeclarationType.Function);
            var endpoints = MoveEndpoints.StdToStd;
            var source =
$@"
Option Explicit

{udtAccessibility} Type MyTestType
    Fizz As Long
    Bizz As String
End Type

Public mFizzBar As MyTestType

Public Function FizzMath(arg1 As Long) As Long
    FooMath = arg1 * mFizzBar.Fizz
End Function

Public Function ConcatBar(arg1 As String) As String
    ConcatBar = arg1 & mFizzBar.Bar
End Function";

            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source);

            StringAssert.Contains($"Public mFizzBar As MyTestType", refactoredCode.Source);
            StringAssert.Contains($"arg1 & mFizzBar.Bar", refactoredCode.Source);
            StringAssert.DoesNotContain("Public Function FizzMath", refactoredCode.Source);

            StringAssert.Contains($"arg1 * {endpoints.SourceModuleName()}.mFizzBar.Fizz", refactoredCode.Destination);
            StringAssert.DoesNotContain($"Public mFizzBar As MyTestType", refactoredCode.Destination);
        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MoveSinglePropertyReferencingPrivateUDTMember_NoStrategy()
        {
            var memberToMove = ("TestValue", DeclarationType.PropertyGet);
            var endpoints = MoveEndpoints.StdToStd;
            var source =
$@"
Option Explicit

Private Type TModuleSource
    TestValue As Long
    Test2Value As String
End Type

Private this As TModuleSource

Public Property Get TestValue() As Long
    TestValue = this.TestValue
End Property

Public Property Let TestValue(ByVal value As Long)
    this.TestValue = value
End Property

Public Property Get Test2Value() As String
    Test2Value = this.Test2Value
End Property

Public Property Let Test2Value(ByVal value As String)
    this.Test2Value = value
End Property
";
            ExecuteSingleTargetMoveThrowsExceptionTest(memberToMove, endpoints, source);
        }

        [TestCase(MoveEndpoints.StdToStd)]
        [TestCase(MoveEndpoints.ClassToStd)]
        [TestCase(MoveEndpoints.FormToStd)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MoveAllPropertiesReferencingPrivateUDT(MoveEndpoints endpoints)
        {
            var memberToMove = ("TestValue", DeclarationType.PropertyGet);
            var source =
$@"
Option Explicit

Private Type TModuleSource
    TestValue As Long
    Test2Value As String
End Type

Private this As TModuleSource

Public Property Get TestValue() As Long
    TestValue = this.TestValue
End Property

Public Property Let TestValue(ByVal value As Long)
    this.TestValue = value
End Property

Public Property Get Test2Value() As String
    Test2Value = this.Test2Value
End Property

Public Property Let Test2Value(ByVal value As String)
    this.Test2Value = value
End Property
";
            MoveMemberModel ModelAdjustment(MoveMemberModel model)
            {
                model.MoveableMemberSetByName("Test2Value").IsSelected = true;
                model.ChangeDestination(endpoints.DestinationModuleName(), ComponentType.StandardModule);
                return model;
            }

            var refactoredCode = RefactorTargets(memberToMove, endpoints, source, string.Empty, ModelAdjustment);

            StringAssert.AreEqualIgnoringCase("Option Explicit", refactoredCode.Source.Trim());

            StringAssert.Contains("Get Test2Value", refactoredCode.Destination);
            StringAssert.Contains("Get Test2Value", refactoredCode.Destination);
            StringAssert.Contains("Let Test2Value", refactoredCode.Destination);
            StringAssert.Contains("Get TestValue", refactoredCode.Destination);
            StringAssert.Contains("Private Type TModuleSource", refactoredCode.Destination);
            StringAssert.DoesNotContain(" Const", refactoredCode.Destination);
            StringAssert.Contains("Private this As TModuleSource", refactoredCode.Destination);
            Assert.IsTrue("Private this As TModuleSource".OccursOnce(refactoredCode.Destination));
        }

        [TestCase("MoveThisUsesProperty", nameof(MoveMemberToStdModule))]
        [TestCase("MoveThisUsesMemberAccess", null)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ProcedureReferencesPropertiesUsingPrivateUDT(string memberTarget, string expectedStrategy)
        {
            var memberToMove = (memberTarget, DeclarationType.Procedure);
            var endpoints = MoveEndpoints.StdToStd;
            var source =
$@"
Option Explicit

Private Type TModuleSource
    TestValue As Long
    Test2Value As String
End Type

Private this As TModuleSource

Public Property Get TestValue() As Long
    TestValue = this.TestValue
End Property

Public Property Let TestValue(ByVal value As Long)
    this.TestValue = value
End Property

Public Property Get Test2Value() As String
    Test2Value = this.Test2Value
End Property

Public Property Let Test2Value(ByVal value As String)
    this.Test2Value = value
End Property

Public Sub MoveThisUsesProperty(arg1 As Long, arg2 As String)
    TestValue = arg1
    Test2Value = arg2
End Sub

Public Sub MoveThisUsesMemberAccess(arg1 As Long, arg2 As String)
    this.TestValue = arg1
    this.Test2Value = arg2
End Sub
";

            if (expectedStrategy is null)
            {
                ExecuteSingleTargetMoveThrowsExceptionTest(memberToMove, endpoints, source);
                return;
            }

            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source);

            StringAssert.DoesNotContain($"Sub {memberTarget}(arg1 As Long, arg2 As String)", refactoredCode.Source);
            StringAssert.Contains("Test2Value", refactoredCode.Source);
            StringAssert.Contains("TestValue", refactoredCode.Source);

            StringAssert.Contains($"Sub {memberTarget}(arg1 As Long, arg2 As String)", refactoredCode.Destination);
            StringAssert.Contains($"{endpoints.SourceModuleName()}.Test2Value", refactoredCode.Destination);
            StringAssert.Contains($"{endpoints.SourceModuleName()}.TestValue", refactoredCode.Destination);
        }

        [TestCase(MoveEndpoints.StdToStd, "Private")]
        [TestCase(MoveEndpoints.StdToStd, "Public")]
        [TestCase(MoveEndpoints.ClassToStd, "Private")]
        [TestCase(MoveEndpoints.FormToStd, "Private")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void FunctionUsesPublicUDTLocally(MoveEndpoints endpoints, string udtAccessibility)
        {
            var memberToMove = ("DoubleTheValue", DeclarationType.Function);
            var source =
$@"
Option Explicit

{udtAccessibility} Type KeyValuePair
    Key As String
    Value As Long
End Type

Public Function DoubleTheValue(key As String, value As Long) As Long
    Dim newKVP As KeyValuePair
    newKVP.Key = key
    newKVP.Value = 2 * value
    DoubleTheValue = newKVP.Value
End Function
";
            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source);
            StringAssert.Contains("Function DoubleTheValue(", refactoredCode.Destination);

            if (udtAccessibility.Equals("Private"))
            {
                StringAssert.AreEqualIgnoringCase("Option Explicit", refactoredCode.Source.Trim());
                StringAssert.Contains("Private Type KeyValuePair", refactoredCode.Destination);
            }
            else
            {
                StringAssert.Contains("Public Type KeyValuePair", refactoredCode.Source);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MovePrivateUDTWithFunction()
        {
            var memberToMove = ("UsePvtType", DeclarationType.Function);
            var endpoints = MoveEndpoints.StdToStd;
            var source =
$@"
Option Explicit

Private Type TestType
    FirstValue As Long
End Type

Private mTestType As TestType

Private Function UsePvtType(arg As Long) As TestType
        mTestType.FirstValue = arg
End Function
";
            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source);

            StringAssert.AreEqualIgnoringCase("Option Explicit", refactoredCode.Source.Trim());

            StringAssert.Contains("Private Type TestType", refactoredCode.Destination);
            StringAssert.Contains("FirstValue As Long", refactoredCode.Destination);
            StringAssert.Contains("Private mTestType As TestType", refactoredCode.Destination);
            StringAssert.Contains("Private Function UsePvtType(arg As Long) As TestType", refactoredCode.Destination);
            Assert.IsTrue("Private Type TestType".OccursOnce(refactoredCode.Destination));
        }
    }
}
