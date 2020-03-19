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
    public class MoveMemberEnumTests : MoveMemberRefactoringActionTestSupportBase
    {
        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void PublicEnumMoveFieldHasReferences()
        {
            var source =
$@"
Option Explicit

Public Enum MyTestEnum
    ValueOne
    ValueTwo
End Enum

Public eFoo As MyTestEnum

Public Function FooMath(arg1 As Long) As Long
    FooMath = arg1 * eFoo
End Function
";

            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, ("eFoo", DeclarationType.Variable), sourceContent: source);

            var callSiteModuleName = "CallSiteModule";

            var otherModuleReference =
    $@"
Option Explicit

Public Function MemberAccessFoo(arg1 As Long) As Long
    MemberAccessFoo = arg1 + {moveDefinition.SourceModuleName}.eFoo * 2
End Function

Public Function WithMemberAccessFoo(arg1 As Long) As Long
    Dim result As Long
    With {moveDefinition.SourceModuleName}
        result = (.eFoo + arg1) * 2
    End With
    WithMemberAccessFoo = result
End Function

Public Function NonQualifiedFoo(arg1 As Long) As Long
    NonQualifiedFoo = (eFoo + arg1) * 2
End Function
";
            moveDefinition.Add(new ModuleDefinition(callSiteModuleName, ComponentType.StandardModule, otherModuleReference));


            var refactoredCode = ExecuteTest(moveDefinition);

            var destinationDeclaration = "Public eFoo As MyTestEnum";

            StringAssert.DoesNotContain("eFoo As MyTestEnum", refactoredCode.Source);
            StringAssert.Contains(destinationDeclaration, refactoredCode.Destination);
            StringAssert.Contains($"{moveDefinition.DestinationModuleName}.eFoo", refactoredCode.Source);
            StringAssert.Contains($"{moveDefinition.DestinationModuleName}.{moveDefinition.SelectedElement}", refactoredCode[callSiteModuleName]);
            StringAssert.Contains($"result = ({moveDefinition.DestinationModuleName}.eFoo + arg1) * 2", refactoredCode[callSiteModuleName]);
            StringAssert.Contains($"NonQualifiedFoo = ({moveDefinition.DestinationModuleName}.eFoo + arg1) * 2", refactoredCode[callSiteModuleName]);
        }

        [TestCase(MoveEndpoints.FormToStd)]
        [TestCase(MoveEndpoints.ClassToStd)]
        [TestCase(MoveEndpoints.StdToStd)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void PrivateEnumMoveField_NoStrategyFound(MoveEndpoints endpoints)
        {
            var source =
$@"
Option Explicit

Private Enum MyTestEnum
    ValueOne
    ValueTwo
End Enum

Private eFoo As MyTestEnum

Public Function FooMath(arg1 As Long) As Long
    FooMath = arg1 * eFoo
End Function
";
            var sourceTuple = MoveMemberTestSupport.EndpointToSourceTuple(endpoints, source);
            var destinationTuple = MoveMemberTestSupport.EndpointToDestinationTuple(endpoints, string.Empty);
            var resultCount = MoveMemberTestSupport.ParseAndTest(ThisTest, sourceTuple, destinationTuple);
            Assert.AreEqual(0, resultCount);

            long ThisTest(RubberduckParserState state, IVBE vbe, IRewritingManager rewritingManager)
            {
                var strategies = MoveMemberTestSupport.RetrieveStrategies(state, "eFoo", DeclarationType.Variable, rewritingManager);
                return strategies.Count();
            }
        }

        [TestCase(MoveEndpoints.StdToStd)]
        [TestCase(MoveEndpoints.ClassToStd)]
        [TestCase(MoveEndpoints.FormToStd)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MoveAllPropertiesReferencingPrivateEnum(MoveEndpoints endpoints)
        {
            var member = ("TestValue", DeclarationType.PropertyGet);
            var source =
$@"
Option Explicit

Private Enum ETestValues
    TestValue
    Test2Value
End Enum

Private mTestValue As ETestValues
Private mTestValue2 As ETestValues

Public Property Get TestValue() As Long
    TestValue = mTestValue
End Property

Public Property Let TestValue(ByVal value As Long)
    mTestValue = value
End Property

Public Property Get Test2Value() As Long
    Test2Value = mTestValue2
End Property

Public Property Let Test2Value(ByVal value As Long)
    mTestValue2 = value
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
            StringAssert.Contains("Private Enum ETestValues", refactoredCode.Destination);
            StringAssert.Contains("Private mTestValue As ETestValues", refactoredCode.Destination);
            StringAssert.Contains("Private mTestValue2 As ETestValues", refactoredCode.Destination);
        }



        [TestCase("MoveThisUsesProperty", nameof(MoveMemberToStdModule))]
        [TestCase("MoveThisUsesMemberAccess", null)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ProcedureReferencesPropertiesUsingPrivateEnum(string memberToMove, string expectedStrategy)
        {
            var member = (memberToMove, DeclarationType.Procedure);
            var source =
$@"
Option Explicit

Private Enum EValues
    NumberOne
    NumberTwo
End Enum

Private mValue As EValues
Private mValue2 As EValues

Public Property Get TestValue() As Long
    TestValue = mValue
End Property

Public Property Let TestValue(ByVal value As Long)
    mValue = value
End Property

Public Property Get Test2Value() As Long
    Test2Value = mValue2
End Property

Public Property Let Test2Value(ByVal value As Long)
    mValue2 = value
End Property

Public Sub MoveThisUsesProperty(arg1 As Long, arg2 As Long)
    TestValue = arg1
    Test2Value = arg2
End Sub

Public Sub MoveThisUsesMemberAccess(arg1 As Long, arg2 As Long)
    mValue = arg1
    mValue2 = arg2
End Sub
";

            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, member, source);

            var refactoredCode = ExecuteTest(moveDefinition);

            if (expectedStrategy is null)
            {
                Assert.IsNull(refactoredCode.StrategyName);
                return;
            }

            StringAssert.DoesNotContain($"Sub {memberToMove}(arg1 As Long, arg2 As Long)", refactoredCode.Source);
            StringAssert.Contains("Test2Value", refactoredCode.Source);
            StringAssert.Contains("TestValue", refactoredCode.Source);

            StringAssert.Contains($"Sub {memberToMove}(arg1 As Long, arg2 As Long)", refactoredCode.Destination);
            StringAssert.Contains($"{moveDefinition.SourceModuleName}.Test2Value", refactoredCode.Destination);
            StringAssert.Contains($"{moveDefinition.SourceModuleName}.TestValue", refactoredCode.Destination);
        }

        [TestCase(MoveEndpoints.StdToStd, "Private")]
        [TestCase(MoveEndpoints.StdToStd, "Public")]
        [TestCase(MoveEndpoints.ClassToStd, "Private")]
        [TestCase(MoveEndpoints.FormToStd, "Private")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void FunctionUsesPublicEnumLocally(MoveEndpoints endpoint, string enumAccessibility)
        {
            var member = ("BoundTheValue", DeclarationType.Function);
            var source =
$@"
Option Explicit

{enumAccessibility} Enum KeyValues
    KeyOne
    KeyTwo
End Enum

Public Function BoundTheValue(key As Long) As Long
    Dim kv As KeyValues
    BoundTheValue = key
    If key > kv.KeyOne Then
        BoundTheValue = kv.KeyOne
    End if
End Function
";

            var moveDefinition = new TestMoveDefinition(endpoint, member, source);

            var refactoredCode = ExecuteTest(moveDefinition);

            StringAssert.Contains("Function BoundTheValue(", refactoredCode.Destination);

            if (enumAccessibility.Equals("Private"))
            {
                StringAssert.AreEqualIgnoringCase("Option Explicit", refactoredCode.Source.Trim());
                StringAssert.Contains("Private Enum KeyValues", refactoredCode.Destination);
            }
            else
            {
                StringAssert.Contains("Public Enum KeyValues", refactoredCode.Source);
            }
        }
    }
}
