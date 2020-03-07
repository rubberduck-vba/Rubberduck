using NUnit.Framework;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.Refactorings.MoveMember;
using RubberduckTests.Mocks;
using System.Linq;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System.Collections.Generic;
using System;

namespace RubberduckTests.Refactoring.MoveMember
{
    [TestFixture]
    public class MoveFieldsOrConstantsTests : MoveMemberRefactoringActionTestSupportBase
    {
        [TestCase(MoveEndpoints.FormToStd, "Private")]
        [TestCase(MoveEndpoints.ClassToStd, "Private")]
        [TestCase(MoveEndpoints.StdToStd, "Public")]
        [TestCase(MoveEndpoints.StdToStd, "Private")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MoveConstant_HasReferences(MoveEndpoints moveEndpoints, string accessibility)
        {
            var source =
$@"
Option Explicit

{accessibility} Const mFoo As Long = 10

Public Function FooMath(arg1 As Long) As Long
    FooMath = arg1 * mFoo * 2
End Function
";

            var moveDefinition = new TestMoveDefinition(moveEndpoints, ("mFoo", DeclarationType.Constant), sourceContent: source);

            var callSiteModuleName = "CallSiteModule";
            if (moveDefinition.IsStdModuleSource && accessibility.Equals(Tokens.Public))
            {
                var otherModuleReference =
    $@"
Option Explicit

Public Function MemberAccessFoo(arg1 As Long) As Long
    MemberAccessFoo = arg1 * {moveDefinition.SourceModuleName}.mFoo * 2
End Function

Public Function WithMemberAccessFoo(arg1 As Long) As Long
    Dim result As Long
    With {moveDefinition.SourceModuleName}
        result = (.mFoo + arg1) * 2
    End With
    WithMemberAccessFoo = result
End Function

Public Function NonQualifiedFoo(arg1 As Long) As Long
    NonQualifiedFoo = (mFoo + arg1) * 2
End Function
";
                moveDefinition.Add(new ModuleDefinition(callSiteModuleName, ComponentType.StandardModule, otherModuleReference));
            }

            var destinationDeclaration = "Public Const mFoo As Long = 10";

            
            var refactoredCode = ExecuteTest(moveDefinition);

            StringAssert.DoesNotContain("Const mFoo As Long = 10", refactoredCode.Source);
            StringAssert.Contains(destinationDeclaration, refactoredCode.Destination);
            StringAssert.Contains($"{moveDefinition.DestinationModuleName}.{moveDefinition.SelectedElement}", refactoredCode.Source);
            if (moveDefinition.IsStdModuleSource && accessibility.Equals(Tokens.Public))
            {
                StringAssert.Contains($"{moveDefinition.DestinationModuleName}.{moveDefinition.SelectedElement}", refactoredCode[callSiteModuleName]);
                StringAssert.Contains($"result = ({moveDefinition.DestinationModuleName}.{moveDefinition.SelectedElement} + arg1) * 2", refactoredCode[callSiteModuleName]);
                StringAssert.Contains($"NonQualifiedFoo = ({moveDefinition.DestinationModuleName}.{moveDefinition.SelectedElement} + arg1) * 2", refactoredCode[callSiteModuleName]);
            }
        }

        [TestCase(MoveEndpoints.FormToStd, "Private", null)]
        [TestCase(MoveEndpoints.ClassToStd, "Private", null)]
        [TestCase(MoveEndpoints.StdToStd, "Public", nameof(MoveMemberToStdModule))]
        [TestCase(MoveEndpoints.StdToStd, "Private", null)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MoveNonAggregateValueTypeField_HasReferences(MoveEndpoints moveEndpoints, string accessibility, string expectedStrategyName)
        {
            var source =
$@"
Option Explicit

{accessibility} mFoo As Long

Public Function FooMath(arg1 As Long) As Long
    FooMath = arg1 * mFoo * 2
End Function
";

            var moveDefinition = new TestMoveDefinition(moveEndpoints, ("mFoo", DeclarationType.Variable), sourceContent: source);

            var callSiteModuleName = "CallSiteModule";
            if (moveDefinition.IsStdModuleSource && accessibility.Equals(Tokens.Public))
            {
                var otherModuleReference =
    $@"
Option Explicit

Public Function MemberAccessFoo(arg1 As Long) As Long
    MemberAccessFoo = arg1 * {moveDefinition.SourceModuleName}.mFoo * 2
End Function

Public Function WithMemberAccessFoo(arg1 As Long) As Long
    Dim result As Long
    With {moveDefinition.SourceModuleName}
        result = (.mFoo + arg1) * 2
    End With
    WithMemberAccessFoo = result
End Function

Public Function NonQualifiedFoo(arg1 As Long) As Long
    NonQualifiedFoo = (mFoo + arg1) * 2
End Function
";
                moveDefinition.Add(new ModuleDefinition(callSiteModuleName, ComponentType.StandardModule, otherModuleReference));
            }

            
            var refactoredCode = ExecuteTest(moveDefinition);
            StringAssert.AreEqualIgnoringCase(expectedStrategyName, refactoredCode.StrategyName);

            if (expectedStrategyName is null)
            {
                return;
            }

            var destinationDeclaration = "Public mFoo As Long";

            StringAssert.DoesNotContain("mFoo As Long", refactoredCode.Source);
            StringAssert.Contains(destinationDeclaration, refactoredCode.Destination);
            StringAssert.Contains($"{moveDefinition.DestinationModuleName}.{moveDefinition.SelectedElement}", refactoredCode.Source);
            StringAssert.Contains($"{moveDefinition.DestinationModuleName}.{moveDefinition.SelectedElement}", refactoredCode[callSiteModuleName]);
            StringAssert.Contains($"result = ({moveDefinition.DestinationModuleName}.{moveDefinition.SelectedElement} + arg1) * 2", refactoredCode[callSiteModuleName]);
            StringAssert.Contains($"NonQualifiedFoo = ({moveDefinition.DestinationModuleName}.{moveDefinition.SelectedElement} + arg1) * 2", refactoredCode[callSiteModuleName]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MoveNonAggregateValueTypeFields_HasReferences()
        {
            var source =
$@"
Option Explicit

Public mFoo As Long
Public mFoo1 As Long
Public mFoo2 As Long
Public mFoo3 As Long

Public Function FooMath(arg1 As Long) As Long
    FooMath = arg1 * mFoo * mFoo2
End Function
";

            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, ("mFoo", DeclarationType.Variable), sourceContent: source);

            var callSiteModuleName = "CallSiteModule";
            var otherModuleReference =
    $@"
Option Explicit

Public Function MemberAccessFoo(arg1 As Long) As Long
    MemberAccessFoo = arg1 * {moveDefinition.SourceModuleName}.mFoo * 2
End Function

Public Function WithMemberAccessFoo(arg1 As Long) As Long
    Dim result As Long
    With {moveDefinition.SourceModuleName}
        result = (.mFoo + arg1) * 2
    End With
    WithMemberAccessFoo = result
End Function

Public Function NonQualifiedFoo(arg1 As Long) As Long
    NonQualifiedFoo = (mFoo2 + arg1) * 2
End Function
";
            moveDefinition.Add(new ModuleDefinition(callSiteModuleName, ComponentType.StandardModule, otherModuleReference));

            var decType = DeclarationType.Variable;
            moveDefinition.SetEndpointContent(source);
            moveDefinition.AddSelectedDeclaration("mFoo1", decType);
            moveDefinition.AddSelectedDeclaration("mFoo2", decType);
            moveDefinition.AddSelectedDeclaration("mFoo3", decType);
            var refactoredCode = ExecuteTest(moveDefinition);

            var destinationDeclaration = "Public mFoo As Long";

            StringAssert.DoesNotContain("mFoo As Long", refactoredCode.Source);
            StringAssert.DoesNotContain("mFoo1 As Long", refactoredCode.Source);
            StringAssert.DoesNotContain("mFoo2 As Long", refactoredCode.Source);
            StringAssert.DoesNotContain("mFoo3 As Long", refactoredCode.Source);
            StringAssert.Contains(destinationDeclaration, refactoredCode.Destination);
            StringAssert.Contains($"{moveDefinition.DestinationModuleName}.{moveDefinition.SelectedElement}", refactoredCode.Source);
            StringAssert.Contains($"{moveDefinition.DestinationModuleName}.mFoo2", refactoredCode.Source);
            StringAssert.Contains($"{moveDefinition.DestinationModuleName}.{moveDefinition.SelectedElement}", refactoredCode[callSiteModuleName]);
            StringAssert.Contains($"result = ({moveDefinition.DestinationModuleName}.{moveDefinition.SelectedElement} + arg1) * 2", refactoredCode[callSiteModuleName]);
            StringAssert.Contains($"NonQualifiedFoo = ({moveDefinition.DestinationModuleName}.mFoo2 + arg1) * 2", refactoredCode[callSiteModuleName]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MoveNonAggregateValueTypeFieldsList_HasReference()
        {
            var source =
$@"
Option Explicit

Public mFoo As Long, mFoo1 As Long, mFoo2 As Long, mFoo3 As Long

Public Function FooMath(arg1 As Long) As Long
    FooMath = arg1 * mFoo * mFoo2
End Function
";

            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, ("mFoo", DeclarationType.Variable), sourceContent: source);

            var callSiteModuleName = "CallSiteModule";
            var otherModuleReference =
    $@"
Option Explicit

Public Function MemberAccessFoo(arg1 As Long) As Long
    MemberAccessFoo = arg1 * {moveDefinition.SourceModuleName}.mFoo * 2
End Function

Public Function WithMemberAccessFoo(arg1 As Long) As Long
    Dim result As Long
    With {moveDefinition.SourceModuleName}
        result = (.mFoo + arg1) * 2
    End With
    WithMemberAccessFoo = result
End Function

Public Function NonQualifiedFoo(arg1 As Long) As Long
    NonQualifiedFoo = (mFoo2 + arg1) * 2
End Function
";
            moveDefinition.Add(new ModuleDefinition(callSiteModuleName, ComponentType.StandardModule, otherModuleReference));

            var decType = DeclarationType.Variable;
            moveDefinition.SetEndpointContent(source);
            moveDefinition.AddSelectedDeclaration("mFoo1", decType);
            moveDefinition.AddSelectedDeclaration("mFoo2", decType);
            moveDefinition.AddSelectedDeclaration("mFoo3", decType);
            var refactoredCode = ExecuteTest(moveDefinition);

            Assert.IsTrue(MoveMemberTestSupport.OccursOnce("Public", refactoredCode.Source));
            StringAssert.DoesNotContain("mFoo As Long", refactoredCode.Source);
            StringAssert.DoesNotContain("mFoo1 As Long", refactoredCode.Source);
            StringAssert.DoesNotContain("mFoo2 As Long", refactoredCode.Source);
            StringAssert.DoesNotContain("mFoo3 As Long", refactoredCode.Source);
            StringAssert.Contains("Public mFoo As Long", refactoredCode.Destination);
            StringAssert.Contains($"Public mFoo1 As Long", refactoredCode.Destination);
            StringAssert.Contains("Public mFoo2 As Long", refactoredCode.Destination);
            StringAssert.Contains($"Public mFoo3 As Long", refactoredCode.Destination);
            StringAssert.Contains($"{moveDefinition.DestinationModuleName}.{moveDefinition.SelectedElement}", refactoredCode.Source);
            StringAssert.Contains($"{moveDefinition.DestinationModuleName}.mFoo2", refactoredCode.Source);
            StringAssert.Contains($"{moveDefinition.DestinationModuleName}.{moveDefinition.SelectedElement}", refactoredCode[callSiteModuleName]);
            StringAssert.Contains($"result = ({moveDefinition.DestinationModuleName}.{moveDefinition.SelectedElement} + arg1) * 2", refactoredCode[callSiteModuleName]);
            StringAssert.Contains($"NonQualifiedFoo = ({moveDefinition.DestinationModuleName}.mFoo2 + arg1) * 2", refactoredCode[callSiteModuleName]);
        }

        [TestCase(MoveEndpoints.FormToStd, "Private")]
        [TestCase(MoveEndpoints.ClassToStd, "Private")]
        [TestCase(MoveEndpoints.StdToStd, "Public")]
        [TestCase(MoveEndpoints.StdToStd, "Private")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MoveNConstantsList_HasReferences(MoveEndpoints moveEndpoints, string accessibility)
        {
            var source =
$@"
Option Explicit

{accessibility} Const mFoo As Long = 0, mFoo1 As Long = 10, mFoo2 As Long = 20, mFoo3 As Long = 30

Public Function FooMath(arg1 As Long) As Long
    FooMath = arg1 * mFoo * mFoo2
End Function
";

            var moveDefinition = new TestMoveDefinition(moveEndpoints, ("mFoo", DeclarationType.Constant), sourceContent: source);

            var callSiteModuleName = "CallSiteModule";
            if (moveDefinition.IsStdModuleSource && accessibility.Equals(Tokens.Public))
            {
                var otherModuleReference =
    $@"
Option Explicit

Public Function MemberAccessFoo(arg1 As Long) As Long
    MemberAccessFoo = arg1 * {moveDefinition.SourceModuleName}.mFoo * 2
End Function

Public Function WithMemberAccessFoo(arg1 As Long) As Long
    Dim result As Long
    With {moveDefinition.SourceModuleName}
        result = (.mFoo + arg1) * 2
    End With
    WithMemberAccessFoo = result
End Function

Public Function NonQualifiedFoo(arg1 As Long) As Long
    NonQualifiedFoo = (mFoo2 + arg1) * 2
End Function
";
                moveDefinition.Add(new ModuleDefinition(callSiteModuleName, ComponentType.StandardModule, otherModuleReference));
            }
            var decType = DeclarationType.Constant;
            moveDefinition.SetEndpointContent(source);
            moveDefinition.AddSelectedDeclaration("mFoo1", decType);
            moveDefinition.AddSelectedDeclaration("mFoo2", decType);
            moveDefinition.AddSelectedDeclaration("mFoo3", decType);
            var refactoredCode = ExecuteTest(moveDefinition);

            StringAssert.DoesNotContain($"{accessibility} Const", refactoredCode.Source);
            StringAssert.DoesNotContain("mFoo As Long", refactoredCode.Source);
            StringAssert.DoesNotContain("mFoo1 As Long", refactoredCode.Source);
            StringAssert.DoesNotContain("mFoo2 As Long", refactoredCode.Source);
            StringAssert.DoesNotContain("mFoo3 As Long", refactoredCode.Source);
            StringAssert.Contains("Public Const mFoo As Long", refactoredCode.Destination);
            StringAssert.Contains($"{accessibility} Const mFoo1 As Long", refactoredCode.Destination);
            StringAssert.Contains("Public Const mFoo2 As Long", refactoredCode.Destination);
            StringAssert.Contains($"{accessibility} Const mFoo3 As Long", refactoredCode.Destination);
            StringAssert.Contains($"{moveDefinition.DestinationModuleName}.{moveDefinition.SelectedElement}", refactoredCode.Source);
            StringAssert.Contains($"{moveDefinition.DestinationModuleName}.mFoo2", refactoredCode.Source);
            if (moveDefinition.IsStdModuleSource && accessibility.Equals(Tokens.Public))
            {
                StringAssert.Contains($"{moveDefinition.DestinationModuleName}.{moveDefinition.SelectedElement}", refactoredCode[callSiteModuleName]);
                StringAssert.Contains($"result = ({moveDefinition.DestinationModuleName}.{moveDefinition.SelectedElement} + arg1) * 2", refactoredCode[callSiteModuleName]);
                StringAssert.Contains($"NonQualifiedFoo = ({moveDefinition.DestinationModuleName}.mFoo2 + arg1) * 2", refactoredCode[callSiteModuleName]);
            }
        }

        [Test]
        [Ignore("This needs to be addressed")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MoveConstantValueReferencesPrivateConstantWithLocalReference()
        {
            var source =
$@"
Option Explicit

Public Const FIZZ As Long = PVT_VALUE

Private Const PVT_VALUE As Long = 75

Private Function Bizz(arg As Long) As Long
    Bizz = arg + PVT_VALUE
End Function
";

            //var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, ("FIZZ", DeclarationType.Constant), sourceContent: source);
            //var refactoredCode = ExecuteTest(moveDefinition);

            //StringAssert.Contains($"FIZZ", refactoredCode.Source);
            //StringAssert.Contains($"PVT_VALUE", refactoredCode.Source);

            var sourceTuple = MoveMemberTestSupport.EndpointToSourceTuple(MoveEndpoints.StdToStd, source);
            var destinationTuple = MoveMemberTestSupport.EndpointToDestinationTuple(MoveEndpoints.StdToStd, string.Empty);

            var resultCount = MoveMemberTestSupport.ParseAndTest(ThisTest, sourceTuple, destinationTuple);

            Assert.AreEqual(0, resultCount);

            long ThisTest(RubberduckParserState state, IVBE vbe, IRewritingManager rewritingManager)
            {
                var strategies = MoveMemberTestSupport.RetrieveStrategies(state, "FIZZ", DeclarationType.Constant, rewritingManager);
                return strategies.Count();
            }
        }

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

        [TestCase(MoveEndpoints.FormToStd)]
        [TestCase(MoveEndpoints.ClassToStd)]
        [TestCase(MoveEndpoints.StdToStd)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void PrivateUDTMoveField_NoStrategyFound(MoveEndpoints endpoints)
        {
            var source =
$@"
Option Explicit

Private Type MyTestType
    Foo As Long
    Bar As String
End Type

Private mFooBar As MyTestType

Public Function FooMath(arg1 As Long) As Long
    FooMath = arg1 * mFooBar.Foo
End Function

Public Function ConcatBar(arg1 As String) As String
    ConcatBar = arg1 & mFooBar.Bar
End Function";

            var sourceTuple = MoveMemberTestSupport.EndpointToSourceTuple(endpoints, source);
            var destinationTuple = MoveMemberTestSupport.EndpointToDestinationTuple(endpoints, string.Empty );
            var resultCount = MoveMemberTestSupport.ParseAndTest(ThisTest, sourceTuple, destinationTuple);
            Assert.AreEqual(0, resultCount);

            long ThisTest(RubberduckParserState state, IVBE vbe, IRewritingManager rewritingManager)
            {
                var strategies = MoveMemberTestSupport.RetrieveStrategies(state, "mFooBar", DeclarationType.Variable, rewritingManager);
                return strategies.Count();
            }
        }

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

        [TestCase(MoveEndpoints.FormToStd, "Public")]
        [TestCase(MoveEndpoints.FormToStd, "Private")]
        [TestCase(MoveEndpoints.ClassToStd, "Public")]
        [TestCase(MoveEndpoints.ClassToStd, "Private")]
        [TestCase(MoveEndpoints.StdToStd, "Public")]
        [TestCase(MoveEndpoints.StdToStd, "Private")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ObjectField_NoStrategyFound(MoveEndpoints endpoints, string accessibility)
        {
            var memberToMove = "mObj";
            var source =
$@"
Option Explicit

{accessibility} mObj As ObjectClass

Public Function FooMath(arg1 As Long) As Long
    if mObj is Nothing Then
        Set mObj = new ObjectClass
    End if

    FooMath = arg1 * mObj.Value
End Function
";

            var objectClass =
$@"
Option Explicit

Private mValue As Long

Private Sub Class_Initialize()
    mValue = 6
End Sub

Public Property Get Value() As Long
    Value = mValue
End Property
";
            var sourceTuple = MoveMemberTestSupport.EndpointToSourceTuple(endpoints, source);
            var destinationTuple = MoveMemberTestSupport.EndpointToDestinationTuple(endpoints, string.Empty);
            var resultCount = MoveMemberTestSupport.ParseAndTest(ThisTest, sourceTuple, destinationTuple, ("ObjectClass", objectClass, ComponentType.ClassModule));
            Assert.AreEqual(0, resultCount);

            long ThisTest(RubberduckParserState state, IVBE vbe, IRewritingManager rewritingManager)
            {
                var strategies = MoveMemberTestSupport.RetrieveStrategies(state, memberToMove, DeclarationType.Variable, rewritingManager);
                return strategies.Count();
            }
        }

        [Test]
        [Category("MoveMember")]
        public void StdToStdPublicArrayMoveFieldHasReferences()
        {
            var source =
$@"
Option Explicit

Public mArray(5) As Long

Public Function FooMath(arg1 As Long) As Long
    FooMath = arg1 * mArray(2)
End Function
";

            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, ("mArray", DeclarationType.Variable), sourceContent: source);

            var callSiteModuleName = "CallSiteModule";

            var otherModuleReference =
    $@"
Option Explicit

Public Function MemberAccessFoo(arg1 As Long) As Long
    MemberAccessFoo = arg1 + {moveDefinition.SourceModuleName}.mArray(3) * 2
End Function

Public Function WithMemberAccessFoo(arg1 As Long) As Long
    Dim result As Long
    With {moveDefinition.SourceModuleName}
        result = (.mArray(2) + arg1) * 2
    End With
    WithMemberAccessFoo = result
End Function

Public Function NonQualifiedFoo(arg1 As Long) As Long
    NonQualifiedFoo = (mArray(1) + arg1) * 2
End Function
";
            moveDefinition.Add(new ModuleDefinition(callSiteModuleName, ComponentType.StandardModule, otherModuleReference));

            
            var refactoredCode = ExecuteTest(moveDefinition);

            var destinationDeclaration = "Public mArray(5) As Long";

            StringAssert.DoesNotContain("mArray(5) As Long", refactoredCode.Source);
            StringAssert.Contains(destinationDeclaration, refactoredCode.Destination);
            StringAssert.Contains($"{moveDefinition.DestinationModuleName}.mArray(2)", refactoredCode.Source);
            StringAssert.Contains($"{moveDefinition.DestinationModuleName}.{moveDefinition.SelectedElement}(3)", refactoredCode[callSiteModuleName]);
            StringAssert.Contains($"result = ({moveDefinition.DestinationModuleName}.mArray(2) + arg1) * 2", refactoredCode[callSiteModuleName]);
            StringAssert.Contains($"NonQualifiedFoo = ({moveDefinition.DestinationModuleName}.mArray(1) + arg1) * 2", refactoredCode[callSiteModuleName]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void CorrectsFieldNameCollisionInDestination()
        {
            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, ("Goo", DeclarationType.PropertyLet));

            var destinationModuleName = moveDefinition.DestinationModuleName;
            var source =
$@"
Option Explicit

Private mfoo As Long
Private mgoo As Long

Public Function Foo(arg1 As Long) As Long
    mfoo = arg1 * 10
    Foo = mfoo
End Function

Public Property Let Goo(arg1 As Long)
    mgoo = arg1
End Property

Public Property Get Goo() As Long
    Goo = mgoo
End Property
";


            var destination =
$@"
Option Explicit

Private mgoo As Long

Public Function Multiply(arg1 As Long) 
    Multiply = mgoo * arg1
End Function
";

            moveDefinition.SetEndpointContent(source, destination);
            var refactorResults = ExecuteTest(moveDefinition);

            var destinationExpectedContent =
$@"
Option Explicit

Private mgoo As Long

Private mgoo1 As Long

Public Property Let Goo(ByVal arg1 As Long)
    mgoo1 = arg1
End Property

Public Property Get Goo() As Long
    Goo = mgoo1
End Property

Public Function Multiply(arg1 As Long) 
    Multiply = mgoo * arg1
End Function
";
            var expectedLines = destinationExpectedContent.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var line in expectedLines)
            {
                StringAssert.Contains(line, refactorResults.Destination);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void SetsNewFieldNameAtExternalReferences()
        {
            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, ("mfoo", DeclarationType.Variable));

            var destinationModuleName = moveDefinition.DestinationModuleName;
            var source =
$@"
Option Explicit

Public mfoo As Long
";


            var destination =
$@"
Option Explicit

Private mfoo As Long

Public Function Multiply(arg1 As Long) 
    Multiply = mfoo * arg1
End Function
";

            var callSiteModuleName = "Module3";
            var callSiteCode =
    $@"
Option Explicit

Private mBar As Long

Public Sub MemberAccess(arg1 As Long)
    mBar = {moveDefinition.SourceModuleName}.mfoo + arg1
End Sub

Public Sub WithMemberAccess(arg2 As Long)
    With {moveDefinition.SourceModuleName}
        mBar = .mfoo + arg2
    End With
End Sub

Public Sub NonQualified(arg3 As Long)
    mBar = mfoo + arg3
End Sub
";
            moveDefinition.Add(new ModuleDefinition(callSiteModuleName, ComponentType.StandardModule, callSiteCode));

            var expectedCallSiteCode =
    $@"
Option Explicit

Private mBar As Long

Public Sub MemberAccess(arg1 As Long)
    mBar = {moveDefinition.DestinationModuleName}.mfoo1 + arg1
End Sub

Public Sub WithMemberAccess(arg2 As Long)
    With {moveDefinition.SourceModuleName}
        mBar = {moveDefinition.DestinationModuleName}.mfoo1 + arg2
    End With
End Sub

Public Sub NonQualified(arg3 As Long)
    mBar = {moveDefinition.DestinationModuleName}.mfoo1 + arg3
End Sub
";
            moveDefinition.SetEndpointContent(source, destination);
            var refactorResults = ExecuteTest(moveDefinition);

            var expectedLines = expectedCallSiteCode.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var line in expectedLines)
            {
                StringAssert.Contains(line, refactorResults[callSiteModuleName]);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void CorrectsConstantNameCollisionInDestination()
        {
            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, ("Goo", DeclarationType.PropertyLet));

            var destinationModuleName = moveDefinition.DestinationModuleName;
            var source =
$@"
Option Explicit

Private mfoo As Long
Private mgoo As Long
Private mgooX As Long
Private Const multiplier As Long = 10

Public Function Foo(arg1 As Long) As Long
    mfoo = arg1 * 10
    Foo = mfoo
End Function

Public Property Let Goo(arg1 As Long)
    mgoo = arg1
    mgooX = mgoo * multiplier
End Property

Public Property Get Goo() As Long
    Goo = mgoo
End Property
";


            var destination =
$@"
Option Explicit

Private Const multiplier As Long = 2

Public Function Multiply(arg1 As Long) 
    Multiply = arg1 * multiplier
End Function
";

            moveDefinition.SetEndpointContent(source, destination);
            var refactorResults = ExecuteTest(moveDefinition);

            var destinationExpectedContent =
$@"
Option Explicit

Private Const multiplier As Long = 2

Private mgoo As Long
Private mgooX As Long
Private Const multiplier1 As Long = 10

Public Property Let Goo(ByVal arg1 As Long)
    mgoo = arg1
    mgooX = mgoo * multiplier1
End Property

Public Property Get Goo() As Long
    Goo = mgoo
End Property

Public Function Multiply(arg1 As Long) 
    Multiply = arg1 * multiplier
End Function
";
            var expectedLines = destinationExpectedContent.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var line in expectedLines)
            {
                StringAssert.Contains(line, refactorResults.Destination);
            }
        }


        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void SetsNewConstantNameAtExternalReferences()
        {
            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, ("mfoo", DeclarationType.Constant));

            var destinationModuleName = moveDefinition.DestinationModuleName;
            var source =
$@"
Option Explicit

Public Const mfoo As Long = 5
";


            var destination =
$@"
Option Explicit

Private mfoo As Long

Public Function Multiply(arg1 As Long) 
    Multiply = mfoo * arg1
End Function
";

            var callSiteModuleName = "Module3";
            var callSiteCode =
    $@"
Option Explicit

Private mBar As Long

Public Sub MemberAccess(arg1 As Long)
    mBar = {moveDefinition.SourceModuleName}.mfoo + arg1
End Sub

Public Sub WithMemberAccess(arg2 As Long)
    With {moveDefinition.SourceModuleName}
        mBar = .mfoo + arg2
    End With
End Sub

Public Sub NonQualified(arg3 As Long)
    mBar = mfoo + arg3
End Sub
";
            moveDefinition.Add(new ModuleDefinition(callSiteModuleName, ComponentType.StandardModule, callSiteCode));

            var expectedCallSiteCode =
    $@"
Option Explicit

Private mBar As Long

Public Sub MemberAccess(arg1 As Long)
    mBar = {moveDefinition.DestinationModuleName}.mfoo1 + arg1
End Sub

Public Sub WithMemberAccess(arg2 As Long)
    With {moveDefinition.SourceModuleName}
        mBar = {moveDefinition.DestinationModuleName}.mfoo1 + arg2
    End With
End Sub

Public Sub NonQualified(arg3 As Long)
    mBar = {moveDefinition.DestinationModuleName}.mfoo1 + arg3
End Sub
";
            moveDefinition.SetEndpointContent(source, destination);
            var refactorResults = ExecuteTest(moveDefinition);

            var expectedLines = expectedCallSiteCode.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var line in expectedLines)
            {
                StringAssert.Contains(line, refactorResults[callSiteModuleName]);
            }
        }

    }
}
