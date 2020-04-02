using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System.Linq;

namespace RubberduckTests.Refactoring.MoveMember
{
    [TestFixture]
    public class MoveMemberUDTTests : MoveMemberRefactoringActionTestSupportBase
    {
        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void PublicUDTMoveField_HasReferences()
        {
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

            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, ("mFooBar", DeclarationType.Variable), sourceContent: source);

            var callSiteModuleName = "CallSiteModule";

            var otherModuleReference =
    $@"
Option Explicit

Public Function MemberAccessFoo(arg1 As Long) As Long
    MemberAccessFoo = arg1 + {moveDefinition.SourceModuleName}.mFooBar.Foo * 2
End Function

Public Function WithMemberAccessFoo(arg1 As Long) As Long
    Dim result As Long
    With {moveDefinition.SourceModuleName}
        result = (.mFooBar.Foo + arg1) * 2
    End With
    WithMemberAccessFoo = result
End Function

Public Function WithMemberAccessFoo2(arg1 As Long) As Long
    Dim result As Long
    With {moveDefinition.SourceModuleName}.mFooBar
        result2 = (.Foo + arg1) * 2
    End With
    WithMemberAccessFoo2 = result2
End Function

Public Function NonQualifiedFoo(arg1 As Long) As Long
    NonQualifiedFoo = (mFooBar.Foo + arg1) * 2
End Function
";
            moveDefinition.Add(new ModuleDefinition(callSiteModuleName, ComponentType.StandardModule, otherModuleReference));


            var refactoredCode = ExecuteTest(moveDefinition);

            var destinationDeclaration = "Public mFooBar As MyTestType";

            StringAssert.DoesNotContain("mFooBar As MyTestType", refactoredCode.Source);
            StringAssert.Contains(destinationDeclaration, refactoredCode.Destination);
            StringAssert.Contains($"{moveDefinition.DestinationModuleName}.mFooBar.Foo", refactoredCode.Source);
            StringAssert.Contains($"{moveDefinition.DestinationModuleName}.mFooBar.Bar", refactoredCode.Source);
            StringAssert.Contains($"{moveDefinition.DestinationModuleName}.{moveDefinition.SelectedElement}", refactoredCode[callSiteModuleName]);
            StringAssert.Contains($"result = ({moveDefinition.DestinationModuleName}.mFooBar.Foo + arg1) * 2", refactoredCode[callSiteModuleName]);
            StringAssert.Contains($"With {moveDefinition.DestinationModuleName}.mFooBar", refactoredCode[callSiteModuleName]);
            StringAssert.Contains($"NonQualifiedFoo = ({moveDefinition.DestinationModuleName}.mFooBar.Foo + arg1) * 2", refactoredCode[callSiteModuleName]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void PrivateUDTMovePublicField_NoStrategyFound()
        {
            var endpoints = MoveEndpoints.StdToStd;
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

            var sourceTuple = MoveMemberTestSupport.EndpointToSourceTuple(endpoints, source);
            var destinationTuple = MoveMemberTestSupport.EndpointToDestinationTuple(endpoints, string.Empty);
            var resultCount = MoveMemberTestSupport.ParseAndTest(ThisTest, sourceTuple, destinationTuple);
            Assert.AreEqual(0, resultCount);

            long ThisTest(RubberduckParserState state, IVBE vbe, IRewritingManager rewritingManager)
            {
                var strategies = MoveMemberTestSupport.RetrieveStrategies(state, "mFooBar", DeclarationType.Variable, rewritingManager);
                return strategies.Count();
            }
        }

        [TestCase("Public")]
        [TestCase("Private")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void PublicUDTFieldReferencedByMovedFunction(string udtAccessibility)
        {
            var member = ("FizzMath", DeclarationType.Function);
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

            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, member, sourceContent: source);

            var refactoredCode = ExecuteTest(moveDefinition);

            StringAssert.Contains($"Public mFizzBar As MyTestType", refactoredCode.Source);
            StringAssert.Contains($"arg1 & mFizzBar.Bar", refactoredCode.Source);
            StringAssert.DoesNotContain("Public Function FizzMath", refactoredCode.Source);

            StringAssert.Contains($"arg1 * {moveDefinition.SourceModuleName}.mFizzBar.Fizz", refactoredCode.Destination);
            StringAssert.DoesNotContain($"Public mFizzBar As MyTestType", refactoredCode.Destination);
        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MoveSinglePropertyReferencingPrivateUDTMember_NoStrategy()
        {
            var member = ("TestValue", DeclarationType.PropertyGet);
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

            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, member, source);

            var refactoredCode = ExecuteTest(moveDefinition);

            Assert.IsNull(refactoredCode.StrategyName);
        }

        [TestCase(MoveEndpoints.StdToStd)]
        [TestCase(MoveEndpoints.ClassToStd)]
        [TestCase(MoveEndpoints.FormToStd)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MoveAllPropertiesReferencingPrivateUDT(MoveEndpoints endpoints)
        {
            var member = ("TestValue", DeclarationType.PropertyGet);
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

            var moveDefinition = new TestMoveDefinition(endpoints, member, source);
            moveDefinition.AddSelectedDeclaration("Test2Value", DeclarationType.PropertyLet);

            var refactoredCode = ExecuteTest(moveDefinition);

            StringAssert.AreEqualIgnoringCase("Option Explicit", refactoredCode.Source.Trim());

            StringAssert.Contains("Get Test2Value", refactoredCode.Destination);
            StringAssert.Contains("Get Test2Value", refactoredCode.Destination);
            StringAssert.Contains("Let Test2Value", refactoredCode.Destination);
            StringAssert.Contains("Get TestValue", refactoredCode.Destination);
            StringAssert.Contains("Private Type TModuleSource", refactoredCode.Destination);
            StringAssert.DoesNotContain(" Const", refactoredCode.Destination);
            StringAssert.Contains("Private this As TModuleSource", refactoredCode.Destination);
            Assert.IsTrue(MoveMemberTestSupport.OccursOnce("Private this As TModuleSource", refactoredCode.Destination));
        }

        [TestCase("MoveThisUsesProperty", nameof(MoveMemberToStdModule))]
        [TestCase("MoveThisUsesMemberAccess", null)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ProcedureReferencesPropertiesUsingPrivateUDT(string memberToMove, string expectedStrategy)
        {
            var member = (memberToMove, DeclarationType.Procedure);
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

            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, member, source);

            var refactoredCode = ExecuteTest(moveDefinition);

            if (expectedStrategy is null)
            {
                Assert.IsNull(refactoredCode.StrategyName);
                return;
            }

            StringAssert.DoesNotContain($"Sub {memberToMove}(arg1 As Long, arg2 As String)", refactoredCode.Source);
            StringAssert.Contains("Test2Value", refactoredCode.Source);
            StringAssert.Contains("TestValue", refactoredCode.Source);

            StringAssert.Contains($"Sub {memberToMove}(arg1 As Long, arg2 As String)", refactoredCode.Destination);
            StringAssert.Contains($"{moveDefinition.SourceModuleName}.Test2Value", refactoredCode.Destination);
            StringAssert.Contains($"{moveDefinition.SourceModuleName}.TestValue", refactoredCode.Destination);
        }

        [TestCase(MoveEndpoints.StdToStd, "Private")]
        [TestCase(MoveEndpoints.StdToStd, "Public")]
        [TestCase(MoveEndpoints.ClassToStd, "Private")]
        [TestCase(MoveEndpoints.FormToStd, "Private")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void FunctionUsesPublicUDTLocally(MoveEndpoints endpoint, string udtAccessibility)
        {
            var member = ("DoubleTheValue", DeclarationType.Function);
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

            var moveDefinition = new TestMoveDefinition(endpoint, member, source);

            var refactoredCode = ExecuteTest(moveDefinition);

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
            var member = ("UsePvtType", DeclarationType.Function);
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

            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, member, source);

            var refactoredCode = ExecuteTest(moveDefinition);

            StringAssert.AreEqualIgnoringCase("Option Explicit", refactoredCode.Source.Trim());

            StringAssert.Contains("Private Type TestType", refactoredCode.Destination);
            StringAssert.Contains("FirstValue As Long", refactoredCode.Destination);
            StringAssert.Contains("Private mTestType As TestType", refactoredCode.Destination);
            StringAssert.Contains("Private Function UsePvtType(arg As Long) As TestType", refactoredCode.Destination);
            Assert.IsTrue(MoveMemberTestSupport.OccursOnce("Private Type TestType", refactoredCode.Destination));
        }
    }
}
