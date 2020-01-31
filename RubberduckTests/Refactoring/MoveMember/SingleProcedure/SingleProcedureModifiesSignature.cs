using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Support = RubberduckTests.Refactoring.MoveMember.MoveMemberTestSupport;

namespace RubberduckTests.Refactoring.MoveMember
{
    [TestFixture, Ignore("Need to support modifying Sub signature")]
    public class SingleProcedureModifiesSignature : MoveMemberTestsBase
    {
        private const string ThisStrategy = "TBD"; // nameof(SingleProcedureSelected);

        //TODO: Is this unnecessary versus a full refactoring test?
        [TestCase(MoveEndpoints.StdToStd)]
        [TestCase(MoveEndpoints.ClassToStd)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MoveMemberStrategy_AddsByRefParams(MoveEndpoints moveType)
        {
            var memberToMove = "CalculateVolume";
            var source =
$@"
Option Explicit

Private mDiameter As Single
Private mVolume As Single
Private mHeight As Single

Public Sub Initialize(diameter As Single, height As Single)
    mDiameter = diameter
    mHeight = height
End Sub

Public Property Get Volume()
    CalculateVolume
    Volume = mVolume
End Property

Public Property Get Area() As Single
    Area = 3.14 * (mDiameter / 2) ^ 2
End Property

Private Sub CalculateVolume()
    mVolume = Area * mHeight
End Sub

Private Sub AnotherAreaRef()
    mVolume = Area * mHeight
End Sub
";

            var moveDefinition = new TestMoveDefinition(moveType, (memberToMove, DeclarationType.Procedure));

            var refactoredCode = RefactoredCode(moveDefinition, source);

            StringAssert.AreEqualIgnoringCase(ThisStrategy, refactoredCode.StrategyName);

            StringAssert.Contains($"ByVal {Support.PARAM_PREFIX}Area As Single, ByVal {Support.PARAM_PREFIX}mHeight As Single, ByRef {Support.PARAM_PREFIX}mVolume As Single", refactoredCode.Destination);
        }

        [TestCase("Public", "LocalPropGet", MoveEndpoints.StdToStd)]
        [TestCase("Private", "LocalPropGet", MoveEndpoints.StdToStd)]
        [TestCase("Public", "LocalPropGet", MoveEndpoints.StdToClass)]
        [TestCase("Private", "LocalPropGet", MoveEndpoints.StdToClass)]
        [TestCase("Public", "LocalPropGet", MoveEndpoints.ClassToClass)]
        [TestCase("Public", "LocalPropGet", MoveEndpoints.ClassToStd)]
        [TestCase("Public", "LocalFunc", MoveEndpoints.StdToStd)]
        [TestCase("Private", "LocalFunc", MoveEndpoints.StdToStd)]
        [TestCase("Public", "LocalFunc", MoveEndpoints.StdToClass)]
        [TestCase("Private", "LocalFunc", MoveEndpoints.StdToClass)]
        [TestCase("Public", "LocalFunc", MoveEndpoints.ClassToClass)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void UseStateParams_SupportingMember(string accessibility, string referencedMember, MoveEndpoints endpoints)
        {
            var source = $@"
Option Explicit

Public bar As Long

Public mVal As Long

{accessibility} mObj As SomeClass

Public Sub Foo(arg1 As Long)
    bar = bar + arg1 + {referencedMember}
End Sub

Public Sub Goo(arg1 As Long)
    bar = bar + arg1 + {referencedMember}
End Sub

{accessibility} Property Get LocalPropGet() As Long
    LocalPropGet = mVal
End Property

{accessibility} Function LocalFunc() As Long
    LocalFunc = 45
End Function

Public Sub LoadValue(arg As SomeClass)
    Set mObj = arg
End Sub
";
        var someClassContent = $@"

Private mValue As Long

Private Sub {MoveMemberResources.Class_Initialize}()
    mValue = 6
End Sub

Public Property Get Value() As Long
    Value = mValue
End Property
";
            var moveDefinition = new TestMoveDefinition(endpoints, ("Foo", DeclarationType.Procedure));
            moveDefinition.Add(new ModuleDefinition("SomeClass", ComponentType.ClassModule, someClassContent));

            var refactoredCode = RefactoredCode(moveDefinition, source);

            StringAssert.AreEqualIgnoringCase(ThisStrategy, refactoredCode.StrategyName);

            var test = refactoredCode.Source;
        }

        //        [TestCase("Class_Initialize", MoveEndpoints.ClassToStd, true)]
        //        [TestCase("Class_Initialize", MoveEndpoints.StdToClass, true)]
        //        [TestCase("Class_Terminate", MoveEndpoints.ClassToStd, true)]
        //        [TestCase("RefInitialize", MoveEndpoints.ClassToStd, true)]
        //        [TestCase("RefTerminate", MoveEndpoints.ClassToStd, true)]
        //        [TestCase("Class_Initialize", MoveEndpoints.StdToStd, false)]
        //        [TestCase("RefTerminate", MoveEndpoints.StdToStd, false)]
        //        [Category("Refactorings")]
        //        [Category("MoveMember")]
        //        public void MoveMemberStrategy_SingleProcedureSelected_ForwardStateAsArgument_IsLifeCycleHandler(string selectedElement, MoveEndpoints moveType, bool expected)
        //        {
        //            var source =
        //$@"
        //Option Explicit

        //Private mFoo As Long

        //Private Sub Class_Initialize()
        //    mFoo = 5
        //End Sub

        //Private Sub Class_Terminate()
        //    mFoo = 6
        //End Sub

        //Private Sub RefInitialize()
        //    Class_Initialize
        //End Sub

        //Private Sub RefTerminate()
        //    Class_Terminate
        //End Sub
        //";

        //            var moveDefinition = new TestMoveDefinition(moveType, selectedElement, sourceContent: source);

        //            bool ThisTest(RubberduckParserState state, IVBE vbe, IRewritingManager mgr)
        //            {
        //                var model = Support.CreateModelAndDefineMove(vbe, moveDefinition, state, mgr);
        //                return MoveMemberStrategyCommon.MoveDefinitionIncludesLifeCycleHandler(model.CurrentScenario.MoveDefinition, model.CurrentScenario as IProvideMoveDeclarationGroups);
        //            }

        //            var vbeStub = Support.BuildVBEStub(moveDefinition, source);
        //            Assert.AreEqual(expected, Support.ParseAndTest(vbeStub, ThisTest));
        //        }

        //        //TODO: What exactly are we testing here?? <= this comment from the past says all that is needed
        //        [TestCase(MoveEndpoints.ClassToClass, "mDestClass", Support.DEFAULT_SOURCE_CLASS_NAME, Support.DEFAULT_DESTINATION_CLASS_NAME)]
        //        [TestCase(MoveEndpoints.ClassToStd, Support.DEFAULT_DESTINATION_MODULE_NAME, Support.DEFAULT_SOURCE_CLASS_NAME, Support.DEFAULT_DESTINATION_MODULE_NAME)]
        //        [Category("Refactorings")]
        //        [Category("MoveMember")]
        //        public void MoveMemberStrategy_SingleProcedureSelfContained_CheckDestinationSignature(MoveEndpoints moveType, params string[] otherParams)
        //        {
        //            var destinationAccessLExpression = otherParams[0];
        //            var sourceModuleName = otherParams[1];
        //            var destinationModuleName = otherParams[2];

        //            var moveDefinition = new TestMoveDefinition(moveType, "Foo");

        //            var destinationClassDeclareAndInitialize = string.Empty;
        //            if (moveDefinition.IsClassDestination)
        //            {
        //                destinationClassDeclareAndInitialize = Support.ClassVariableDeclareAndInitialize(destinationModuleName, destinationAccessLExpression);
        //            }

        //            var source =
        //$@"
        //Option Explicit

        //Private mfoo As Long
        //Private mgoo As Long

        //{destinationClassDeclareAndInitialize}

        //Public Sub Foo(arg1 As Long)
        //    mfoo = arg1
        //    If {destinationAccessLExpression}.LogIsEnabled Then
        //        {destinationAccessLExpression}.Log ""Foo called""
        //        {destinationAccessLExpression}.Entries = {destinationAccessLExpression}.Entries + 1
        //    Endif
        //End Sub

        //Public Property Let Goo(arg1 As Long)
        //    mgoo = arg1
        //    {destinationAccessLExpression}.Log ""Let Goo called""
        //End Property

        //Public Property Get Goo() As Long
        //    Goo = mgoo
        //    {destinationAccessLExpression}.Log ""Get Goo called""
        //End Property";


        //            var destination =
        //$@"
        //Option Explicit

        //Private Const LOG_IS_ENABLED = True

        //Public Entries As Long

        //Public Property Get LogIsEnabled()
        //    LogIsEnabled = LOG_IS_ENABLED
        //End Property

        //Public Sub Log(msg As String)
        //End Sub";

        //            var vbeStub = Support.BuildVBEStub(moveDefinition, source);

        //            string ThisTest(RubberduckParserState state, IVBE vbe, IRewritingManager mgr)
        //            {

        //                var model = Support.CreateModelAndDefineMove(vbe, moveDefinition, state, mgr);
        //                var moveGroups = model.CurrentScenario as IProvideMoveDeclarationGroups;
        //                var member = moveGroups.SelectedElements.AllDeclarations.Where(m => m.IsMember()).First();
        //                var sut = new MoveMemberContentInfo(model.CurrentScenario as IProvideMoveDeclarationGroups /*.Source*/);
        //                return sut.DestinationSignatureParameters(member);
        //            }

        //            var argSignature = Support.ParseAndTest(vbeStub, ThisTest);

        //            StringAssert.AreEqualIgnoringCase($"arg1 As Long", argSignature);
        //        }

        //Removed because the solution requires changing the moved procedure' signature
        //        [TestCase(MoveEndpoints.StdToStd)]
        //        [TestCase(MoveEndpoints.ClassToStd)]
        //        //[TestCase(MoveEndpoints.StdToClass)]
        //        //[TestCase(MoveEndpoints.ClassToClass)]
        //        [Category("Refactorings")]
        //        [Category("MoveMember")]
        //        public void MoveMemberStrategy_MoveSupportingVariables_HandlesDeclarationLists(MoveEndpoints moveType)
        //        {
        //            //Move Foo - and takes variable1 and variable2 along with it
        //            var memberToMove = ("Foo", DeclarationType.Procedure);
        //            var source =
        //$@"
        //Option Explicit

        //Private variable1 As Long, variable2 As Long, variable3 As Long

        //Public myPublicVariable As Long

        //Public Sub Initialize(v1 As Long, v2 As Long, v3 As Long)
        //    variable1 = v1
        //    variable2 = v2
        //    variable3 = v3
        //End Sub

        //Public Sub Foo()
        //    myPublicVariable = variable2 + variable3
        //End Sub

        //Public Sub Bar(arg1 As Long)
        //    variable1 = variable1 + arg1
        //End Sub";

        //            var moveDefinition = new TestMoveDefinition(moveType, memberToMove);

        //            var refactoredCode = RefactoredCode(moveDefinition, source);

        //            StringAssert.AreEqualIgnoringCase(ThisStrategy, refactoredCode.StrategyName);

        //            var destinationExpectedSignature = $"Public Sub Foo({Tokens.ByVal} {Support.PARAM_PREFIX}variable2 As Long, {Tokens.ByVal} {Support.PARAM_PREFIX}variable3 As Long, {Tokens.ByRef} {Support.PARAM_PREFIX}myPublicVariable As Long)";
        //            var destinationExpectedBody = $"{Support.PARAM_PREFIX}myPublicVariable = {Support.PARAM_PREFIX}variable2 + {Support.PARAM_PREFIX}variable3";
        //            var sourceExpectedForwardingCall = "Foo variable2, variable3, myPublicVariable";

        //            StringAssert.Contains(sourceExpectedForwardingCall, refactoredCode.Source);
        //            StringAssert.Contains(destinationExpectedSignature, refactoredCode.Destination);
        //            StringAssert.Contains(destinationExpectedBody, refactoredCode.Destination);
        //        }

        //REmvoed because this is strategy that moves private constants...skip this for the initial release
        //        [TestCase(MoveEndpoints.StdToStd)]
        //        //[TestCase(MoveEndpoints.StdToClass)]
        //        //[TestCase(MoveEndpoints.ClassToClass)]
        //        [TestCase(MoveEndpoints.ClassToStd)]
        //        //[TestCase(MoveEndpoints.FormToClass)]
        //        [TestCase(MoveEndpoints.FormToStd)]
        //        [Category("Refactorings")]
        //        [Category("MoveMember")]
        //        public void MoveSupportingConstants_PassesRetainedConstantAsArgument(MoveEndpoints endpoints)
        //        {
        //            var source =
        //$@"
        //Option Explicit

        //Public myPublicVariable As Long
        //Private Const constant1 As Long = 1, constant2 As Long = 2, constant3 As Long = 3

        //Public Sub Foo()
        //    myPublicVariable = myPublicVariable + constant2 + constant3
        //End Sub

        //Public Sub Bar(arg1 As Long)
        //    myPublicVariable = myPublicVariable + constant1 + constant2 + arg1
        //End Sub";
        //            var moveDefinition = new TestMoveDefinition(endpoints, ("Foo", DeclarationType.Procedure), "Module1", "Module2");

        //            var refactoredCode = RefactoredCode(moveDefinition, source);

        //            StringAssert.AreEqualIgnoringCase(ThisStrategy, refactoredCode.StrategyName);

        //            var sourceToBeMoved = "Private Const constant3 As Long = 3";
        //            var sourceModifiedConstantListFormatError = "Private constant1 As Long = 1, constant2 As Long = 2"; //the comma should not be there
        //            var destinationExpectedSignature = $"Public Sub Foo(ByVal {Support.PARAM_PREFIX}constant2 As Long, ByRef {Support.PARAM_PREFIX}myPublicVariable As Long)";
        //            var destinationExpectedBody = $"{Support.PARAM_PREFIX}myPublicVariable = {Support.PARAM_PREFIX}myPublicVariable + {Support.PARAM_PREFIX}constant2 + constant3";

        //            StringAssert.DoesNotContain(sourceToBeMoved, refactoredCode.Source);
        //            StringAssert.DoesNotContain(sourceModifiedConstantListFormatError, refactoredCode.Source);
        //            StringAssert.Contains(destinationExpectedSignature, refactoredCode.Destination);
        //            StringAssert.Contains(destinationExpectedBody, refactoredCode.Destination);
        //        }

        //        [Test]
        //        [Category("Refactorings")]
        //        [Category("MoveMember")]
        //        public void RemovesScopeResolutionInDestinationPreview()
        //        {
        //            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, "Foo");
        //            var source =
        //$@"
        //Option Explicit

        //Private mfoo As Long
        //Private mgoo As Long

        //Public Sub Foo(arg1 As Long)
        //    mfoo = arg1
        //    If {moveDefinition.DestinationModuleName}.LOG_ENABLED Then
        //        {moveDefinition.DestinationModuleName}.Log ""Foo called""
        //        {moveDefinition.DestinationModuleName}.Entries = {moveDefinition.DestinationModuleName}.Entries + 1
        //    Endif
        //End Sub

        //Public Property Let Goo(arg1 As Long)
        //    mgoo = arg1
        //    {moveDefinition.DestinationModuleName}.Log ""Let Goo called""
        //End Property

        //Public Property Get Goo() As Long
        //    Goo = mgoo
        //    {moveDefinition.DestinationModuleName}.Log ""Get Goo called""
        //End Property";


        //            var destination =
        //$@"
        //Option Explicit

        //Public Const LOG_ENABLED As Boolean = True

        //Public Entries As Long

        //Public Sub SomeOtherUserOfFoo()
        //    {moveDefinition.SourceModuleName}.Foo 10
        //End Sub

        //Public Sub Log(msg As String)
        //End Sub";

        //            string ThisTest(RubberduckParserState state, IVBE vbe, IRewritingManager mgr)
        //            {
        //                var model = Support.CreateModelAndDefineMove(vbe, moveDefinition, state, mgr);
        //                return model.PreviewDestination();
        //            }

        //            var vbeStub = Support.BuildVBEStub(moveDefinition, source, destination);
        //            var destinationModuleContent = Support.ParseAndTest(vbeStub, ThisTest);

        //            StringAssert.DoesNotContain($"If {moveDefinition.DestinationModuleName}.LOG_ENABLED Then", destinationModuleContent);
        //            StringAssert.DoesNotContain($"{moveDefinition.DestinationModuleName}.Log", destinationModuleContent);
        //            StringAssert.DoesNotContain($"{moveDefinition.DestinationModuleName}.Entries", destinationModuleContent);
        //            StringAssert.Contains("If LOG_ENABLED Then", destinationModuleContent);
        //            StringAssert.Contains("Entries = Entries + 1", destinationModuleContent);
        //            StringAssert.DoesNotContain(".Foo 10", destinationModuleContent);
        //            StringAssert.Contains("Foo 10", destinationModuleContent);
        //        }

        //        [Test]
        //        [Category("Refactorings")]
        //        [Category("MoveMember")]
        //        public void MoveMemberStrategy_StdToStd_Variable_ReferencedByRetainedMembers()
        //        {
        //            var source = $@"
        //Option Explicit

        //Private bar As Long

        //Public Sub Foo(arg1 As Long)
        //    bar = bar + arg1
        //End Sub

        //Public Sub Goo(arg1 As Long)
        //    bar = bar + arg1
        //End Sub";
        //            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, ("Foo", DeclarationType.Procedure));

        //            var refactoredCode = RefactoredCode(moveDefinition, source);

        //            StringAssert.AreEqualIgnoringCase(ThisStrategy, refactoredCode.StrategyName);

        //            var destinationSignature = $"Public Sub Foo(ByRef {Support.PARAM_PREFIX}bar As Long, arg1 As Long)";
        //            var destinationBody = $"{Support.PARAM_PREFIX}bar = {Support.PARAM_PREFIX}bar + arg1";

        //            var sourceForwardingCall = $"{moveDefinition.DestinationModuleName}.Foo bar, arg1";
        //            StringAssert.Contains(sourceForwardingCall, refactoredCode.Source);

        //            StringAssert.Contains(destinationSignature, refactoredCode.Destination);
        //            StringAssert.Contains(destinationBody, refactoredCode.Destination);
        //        }

        //        [TestCase(MoveEndpoints.StdToStd, "Private", "Private")]
        //        [TestCase(MoveEndpoints.StdToStd, "Public", "Private")]
        //        [TestCase(MoveEndpoints.StdToStd, "Private", "Public")]
        //        [TestCase(MoveEndpoints.StdToStd, "Public", "Public")]
        //        [TestCase(MoveEndpoints.ClassToStd, "Private", "Private")]
        //        [TestCase(MoveEndpoints.ClassToStd, "Public", "Private")]
        //        [TestCase(MoveEndpoints.ClassToStd, "Private", "Public")]
        //        [TestCase(MoveEndpoints.ClassToStd, "Public", "Public")]
        //        [Category("Refactorings")]
        //        [Category("MoveMember")]
        //        public void MemberAccessibility_ReturnsStrategy(MoveEndpoints endpoints, string memberAccessibility, string variableAccessibility)
        //        {
        //            var source = $@"
        //Option Explicit

        //{variableAccessibility} bar As Long

        //{memberAccessibility} Sub Foo(arg1 As Long)
        //    bar = bar + arg1
        //End Sub

        //Private Sub UsesFoo()
        //    Foo 5
        //End Sub";

        //            var moveDefinition = new TestMoveDefinition(endpoints, "Foo");

        //            IMoveMemberRefactoringStrategy ThisTest(RubberduckParserState state, IVBE vbe, IRewritingManager mgr)
        //            {
        //                var model = Support.CreateModelAndDefineMove(vbe, moveDefinition, state, mgr);
        //                return model.Strategy;
        //            }

        //            var strategy = Support.ParseAndTest(ThisTest, moveDefinition, source);
        //            Assert.AreEqual(nameof(SingleProcedureSelected), strategy?.GetType().Name ?? null);
        //        }

        //        [TestCase(MoveEndpoints.StdToStd)]
        //        [Category("Refactorings")]
        //        [Category("MoveMember")]
        //        public void MoveMemberStrategy_PrivateVariableReferencedBySupportingMember_Refactors(MoveEndpoints endpoints)
        //        {
        //            var moveDefinition = new TestMoveDefinition(endpoints, selection: ("Foo", DeclarationType.Procedure));

        //            var source = @"
        //Option Explicit

        //Private bar As Long

        //Public Sub Foo(arg1 As Long)
        //    AddValue(arg1)
        //End Sub

        //Public Sub Goo(arg1 As Long)
        //End Sub

        //Private Sub AddValue(arg1 As Long)
        //    bar = bar + arg1
        //End Sub";
        //            var refactoredCode = RefactoredCode(moveDefinition, source);

        //            StringAssert.AreEqualIgnoringCase(ThisStrategy, refactoredCode.StrategyName);

        //            StringAssert.Contains("Public Sub Goo(arg1 As Long)", refactoredCode.Source);
        //            StringAssert.DoesNotContain("Foo", refactoredCode.Source);
        //            StringAssert.DoesNotContain("AddValue(arg1 As Long)", refactoredCode.Source);
        //            StringAssert.DoesNotContain("bar", refactoredCode.Source);

        //            StringAssert.Contains("Public Sub Foo(arg1 As Long)", refactoredCode.Destination);
        //            StringAssert.Contains("AddValue(arg1)", refactoredCode.Destination);
        //            StringAssert.Contains("Private bar As Long", refactoredCode.Destination);
        //            StringAssert.Contains("bar = bar + arg1", refactoredCode.Destination);
        //        }
    }
}
