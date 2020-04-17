using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.VBEditor.SafeComWrappers;

namespace RubberduckTests.Refactoring.MoveMember
{
    [TestFixture]
    public class MoveEnumToStdModuleTests : MoveMemberRefactoringActionTestSupportBase
    {
        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void PublicEnumMoveFieldHasReferences()
        {
            var endpoints = MoveEndpoints.StdToStd;
            var memberToMove = ("eFoo", DeclarationType.Variable);
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

            //var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, ("eFoo", DeclarationType.Variable), sourceContent: source);

            var callSiteModuleName = "CallSiteModule";

            var otherModuleReference =
    $@"
Option Explicit

Public Function MemberAccessFoo(arg1 As Long) As Long
    MemberAccessFoo = arg1 + {endpoints.SourceModuleName()}.eFoo * 2
End Function

Public Function WithMemberAccessFoo(arg1 As Long) As Long
    Dim result As Long
    With {endpoints.SourceModuleName()}
        result = (.eFoo + arg1) * 2
    End With
    WithMemberAccessFoo = result
End Function

Public Function NonQualifiedFoo(arg1 As Long) As Long
    NonQualifiedFoo = (eFoo + arg1) * 2
End Function
";
            //moveDefinition.Add(new ModuleDefinition(callSiteModuleName, ComponentType.StandardModule, otherModuleReference));


            //var refactoredCode = ExecuteTest(moveDefinition, moveDefinition.ModelBuilder);
            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, 
                endpoints.ToSourceTuple(source),
                endpoints.ToDestinationTuple(string.Empty),
                (callSiteModuleName, otherModuleReference, ComponentType.StandardModule));

            var destinationDeclaration = "Public eFoo As MyTestEnum";

            StringAssert.DoesNotContain("eFoo As MyTestEnum", refactoredCode.Source);
            StringAssert.Contains(destinationDeclaration, refactoredCode.Destination);
            StringAssert.Contains($"{endpoints.DestinationModuleName()}.eFoo", refactoredCode.Source);
            StringAssert.Contains($"{endpoints.DestinationModuleName()}.eFoo", refactoredCode[callSiteModuleName]);
            StringAssert.Contains($"result = ({endpoints.DestinationModuleName()}.eFoo + arg1) * 2", refactoredCode[callSiteModuleName]);
            StringAssert.Contains($"NonQualifiedFoo = ({endpoints.DestinationModuleName()}.eFoo + arg1) * 2", refactoredCode[callSiteModuleName]);
        }

        [TestCase(MoveEndpoints.StdToStd)]
        [TestCase(MoveEndpoints.ClassToStd)]
        [TestCase(MoveEndpoints.FormToStd)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MoveAllPropertiesReferencingPrivateEnum(MoveEndpoints endpoints)
        {
            var memberToMove = ("TestValue", DeclarationType.PropertyGet);
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

            //var moveDefinition = new TestMoveDefinition(endpoints, memberToMove, source);
            //moveDefinition.AddSelectedDeclaration("Test2Value", DeclarationType.PropertyLet);
            MoveMemberModel ModelAdjustment(MoveMemberModel model)
            {
                model.MoveableMemberSetByName("Test2Value").IsSelected = true;
                model.ChangeDestination(endpoints.DestinationModuleName(), ComponentType.StandardModule);
                return model;
            }

            //var refactoredCode = ExecuteTest(moveDefinition, moveDefinition.ModelBuilder);
            var refactoredCode = RefactorTargets(memberToMove, endpoints, source, string.Empty, ModelAdjustment);

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
        public void ProcedureReferencesPropertiesUsingPrivateEnum(string identifier, string expectedStrategy)
        {
            var memberToMove = (identifier, DeclarationType.Procedure);
            var endpoints = MoveEndpoints.StdToStd;
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

            //var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, memberToMove, source);

            if (expectedStrategy is null)
            {
                //Assert.Throws<MoveMemberUnsupportedMoveException>(() => ExecuteTest(moveDefinition, moveDefinition.ModelBuilder));
                ExecuteSingleTargetMoveThrowsExceptionTest(memberToMove, endpoints, source);
                return;
            }
            //var refactoredCode = ExecuteTest(moveDefinition, moveDefinition.ModelBuilder);
            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source);

            StringAssert.DoesNotContain($"Sub {identifier}(arg1 As Long, arg2 As Long)", refactoredCode.Source);
            StringAssert.Contains("Test2Value", refactoredCode.Source);
            StringAssert.Contains("TestValue", refactoredCode.Source);

            StringAssert.Contains($"Sub {identifier}(arg1 As Long, arg2 As Long)", refactoredCode.Destination);
            StringAssert.Contains($"{endpoints.SourceModuleName()}.Test2Value", refactoredCode.Destination);
            StringAssert.Contains($"{endpoints.SourceModuleName()}.TestValue", refactoredCode.Destination);
        }

        [TestCase(MoveEndpoints.StdToStd, "Private")]
        [TestCase(MoveEndpoints.StdToStd, "Public")]
        [TestCase(MoveEndpoints.ClassToStd, "Private")]
        [TestCase(MoveEndpoints.FormToStd, "Private")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void FunctionUsesPublicEnumLocally(MoveEndpoints endpoints, string enumAccessibility)
        {
            var memberToMove = ("BoundTheValue", DeclarationType.Function);
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

            //var moveDefinition = new TestMoveDefinition(endpoints, memberToMove, source);

            //var refactoredCode = ExecuteTest(moveDefinition, moveDefinition.ModelBuilder);
            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source);

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


        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MovePrivateEnumWithFunction()
        {
            var memberToMove = ("UsePvtEnum", DeclarationType.Function);
            var endpoints = MoveEndpoints.StdToStd;
            var source =
$@"
Option Explicit

Private Enum KeyValues
    KeyOne
    KeyTwo
End Enum

Private mKV As KeyValues

Private Function UsePvtEnum(arg As Long) As KeyValues
    If arg = KeyOne OR arg = KeyTwo Then mKV = arg

    UsePvtEnum = mKV
End Function
";

            //var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, memberToMove, source);

            //var refactoredCode = ExecuteTest(moveDefinition, moveDefinition.ModelBuilder);

            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source);

            StringAssert.AreEqualIgnoringCase("Option Explicit", refactoredCode.Source.Trim());

            StringAssert.Contains("Private Enum KeyValues", refactoredCode.Destination);
            StringAssert.Contains("Private mKV As KeyValues", refactoredCode.Destination);
            StringAssert.Contains("Private Function UsePvtEnum(arg As Long) As KeyValues", refactoredCode.Destination);
            Assert.IsTrue("Private Enum KeyValues".OccursOnce(refactoredCode.Destination));
        }

        [TestCase("KeyOne", null)]
        [TestCase("KeyValues", nameof(MoveMemberToStdModule))]
        [Category("Refactorings")]
        [Category(nameof(NameConflictFinder))]
        [Category("MoveMember")]
        public void MovePrivateEnumRespectsDestinationNameCollision(string memberName, string expectedStrategy)
        {
            var memberToMove = ("UsePvtEnum", DeclarationType.Function);
            var endpoints = MoveEndpoints.StdToStd;
            var source =
$@"
Option Explicit

Private Enum KeyValues
    KeyOne
    KeyTwo
End Enum

Private mKV As KeyValues

Private Function UsePvtEnum(arg As Long) As KeyValues
    If arg = KeyOne OR arg = KeyTwo Then mKV = arg

    UsePvtEnum = mKV
End Function
";

            var destination =
$@"
Option Explicit

Private Sub {memberName}(arg As Long)
End Sub
";

            //var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, memberToMove, source, destination);

            if (expectedStrategy is null)
            {
                //Assert.Throws<MoveMemberUnsupportedMoveException>(() => ExecuteTest(moveDefinition, moveDefinition.ModelBuilder));
                ExecuteSingleTargetMoveThrowsExceptionTest(memberToMove, endpoints, source, destination);
                return;
            }

            //var refactoredCode = ExecuteTest(moveDefinition, moveDefinition.ModelBuilder);

            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, endpoints.SourceModuleName(), source, destination);
            StringAssert.AreEqualIgnoringCase(expectedStrategy, refactoredCode.StrategyName);
        }
    }
}
