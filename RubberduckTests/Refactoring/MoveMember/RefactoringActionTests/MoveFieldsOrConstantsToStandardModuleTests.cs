using NUnit.Framework;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.Refactorings.MoveMember;
using System;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings;
using System.Collections.Generic;

namespace RubberduckTests.Refactoring.MoveMember
{
    [TestFixture]
    public class MoveFieldsOrConstantsToStandardModuleTests : MoveMemberRefactoringActionTestSupportBase
    {
        [TestCase(MoveEndpoints.FormToStd, "Private")]
        [TestCase(MoveEndpoints.ClassToStd, "Private")]
        [TestCase(MoveEndpoints.StdToStd, "Public")]
        [TestCase(MoveEndpoints.StdToStd, "Private")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MoveConstant_HasReferences(MoveEndpoints endpoints, string accessibility)
        {
            var memberToMove = ("mFoo", DeclarationType.Constant);
            var source =
$@"
Option Explicit

{accessibility} Const mFoo As Long = 10

Public Function FooMath(arg1 As Long) As Long
    FooMath = arg1 * mFoo * 2
End Function
";

            var moduleTuples = new List<(string, string, ComponentType)>();

            var callSiteModuleName = "CallSiteModule";
            if (endpoints.IsStdModuleSource() && accessibility.Equals(Tokens.Public))
            {
                var otherModuleReference =
    $@"
Option Explicit

Public Function MemberAccessFoo(arg1 As Long) As Long
    MemberAccessFoo = arg1 * {endpoints.SourceModuleName()}.mFoo * 2
End Function

Public Function WithMemberAccessFoo(arg1 As Long) As Long
    Dim result As Long
    With {endpoints.SourceModuleName()}
        result = (.mFoo + arg1) * 2
    End With
    WithMemberAccessFoo = result
End Function

Public Function NonQualifiedFoo(arg1 As Long) As Long
    NonQualifiedFoo = (mFoo + arg1) * 2
End Function
";
                moduleTuples.Add((callSiteModuleName, otherModuleReference, ComponentType.StandardModule));
            }

            var destinationDeclaration = "Public Const mFoo As Long = 10";

            moduleTuples.Add(endpoints.ToSourceTuple(source));
            moduleTuples.Add(endpoints.ToDestinationTuple(string.Empty));

            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, moduleTuples.ToArray());

            StringAssert.DoesNotContain("Const mFoo As Long = 10", refactoredCode.Source);
            StringAssert.Contains(destinationDeclaration, refactoredCode.Destination);
            StringAssert.Contains($"{endpoints.DestinationModuleName()}.mFoo", refactoredCode.Source);
            if (endpoints.IsStdModuleSource() && accessibility.Equals(Tokens.Public))
            {
                StringAssert.Contains($"{endpoints.DestinationModuleName()}.mFoo", refactoredCode[callSiteModuleName]);
                StringAssert.Contains($"result = ({endpoints.DestinationModuleName()}.mFoo + arg1) * 2", refactoredCode[callSiteModuleName]);
                StringAssert.Contains($"NonQualifiedFoo = ({endpoints.DestinationModuleName()}.mFoo + arg1) * 2", refactoredCode[callSiteModuleName]);
            }
        }

        [TestCase(MoveEndpoints.FormToStd, "Private", true)]
        [TestCase(MoveEndpoints.ClassToStd, "Private", true)]
        [TestCase(MoveEndpoints.StdToStd, "Public", false)]
        [TestCase(MoveEndpoints.StdToStd, "Private", true)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MoveNonAggregateValueTypeField_HasReferences(MoveEndpoints endpoints, string accessibility, bool throwsException)
        {
            var memberToMove = ("mFoo", DeclarationType.Variable);
            var source =
$@"
Option Explicit

{accessibility} mFoo As Long

Public Function FooMath(arg1 As Long) As Long
    FooMath = arg1 * mFoo * 2
End Function
";
            var moduleTuples = new List<(string, string, ComponentType)>();

            var callSiteModuleName = "CallSiteModule";
            if (endpoints.IsStdModuleSource() && accessibility.Equals(Tokens.Public))
            {
                var otherModuleReference =
    $@"
Option Explicit

Public Function MemberAccessFoo(arg1 As Long) As Long
    MemberAccessFoo = arg1 * {endpoints.SourceModuleName()}.mFoo * 2
End Function

Public Function WithMemberAccessFoo(arg1 As Long) As Long
    Dim result As Long
    With {endpoints.SourceModuleName()}
        result = (.mFoo + arg1) * 2
    End With
    WithMemberAccessFoo = result
End Function

Public Function NonQualifiedFoo(arg1 As Long) As Long
    NonQualifiedFoo = (mFoo + arg1) * 2
End Function
";
                moduleTuples.Add((callSiteModuleName, otherModuleReference, ComponentType.StandardModule));
            }

            moduleTuples.Add(endpoints.ToSourceTuple(source));
            moduleTuples.Add(endpoints.ToDestinationTuple(string.Empty));

            if (throwsException)
            {
                ExecuteSingleTargetMoveThrowsExceptionTest(memberToMove, endpoints, moduleTuples.ToArray());
                return;
            }

            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, moduleTuples.ToArray());

            var destinationDeclaration = "Public mFoo As Long";

            StringAssert.DoesNotContain("mFoo As Long", refactoredCode.Source);
            StringAssert.Contains(destinationDeclaration, refactoredCode.Destination);
            StringAssert.Contains($"{endpoints.DestinationModuleName()}.mFoo", refactoredCode.Source);
            StringAssert.Contains($"{endpoints.DestinationModuleName()}.mFoo", refactoredCode[callSiteModuleName]);
            StringAssert.Contains($"result = ({endpoints.DestinationModuleName()}.mFoo + arg1) * 2", refactoredCode[callSiteModuleName]);
            StringAssert.Contains($"NonQualifiedFoo = ({endpoints.DestinationModuleName()}.mFoo + arg1) * 2", refactoredCode[callSiteModuleName]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MoveNonAggregateValueTypeFields_HasReferences()
        {
            var endpoints = MoveEndpoints.StdToStd;
            var memberToMove = ("mFoo", DeclarationType.Variable);
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
            var callSiteModuleName = "CallSiteModule";
            var otherModuleReference =
    $@"
Option Explicit

Public Function MemberAccessFoo(arg1 As Long) As Long
    MemberAccessFoo = arg1 * {endpoints.SourceModuleName()}.mFoo * 2
End Function

Public Function WithMemberAccessFoo(arg1 As Long) As Long
    Dim result As Long
    With {endpoints.SourceModuleName()}
        result = (.mFoo + arg1) * 2
    End With
    WithMemberAccessFoo = result
End Function

Public Function NonQualifiedFoo(arg1 As Long) As Long
    NonQualifiedFoo = (mFoo2 + arg1) * 2
End Function
";
            MoveMemberModel ModelAdjustment(MoveMemberModel model)
            {
                model.MoveableMemberSetByName("mFoo1").IsSelected = true;
                model.MoveableMemberSetByName("mFoo2").IsSelected = true;
                model.MoveableMemberSetByName("mFoo3").IsSelected = true;
                model.ChangeDestination(endpoints.DestinationModuleName(), ComponentType.StandardModule);
                return model;
            }

            var refactoredCode = RefactorTargets(memberToMove, endpoints, ModelAdjustment,
                endpoints.ToSourceTuple(source),
                endpoints.ToDestinationTuple(string.Empty),
                (callSiteModuleName, otherModuleReference, ComponentType.StandardModule));

            var destinationDeclaration = "Public mFoo As Long";

            StringAssert.DoesNotContain("mFoo As Long", refactoredCode.Source);
            StringAssert.DoesNotContain("mFoo1 As Long", refactoredCode.Source);
            StringAssert.DoesNotContain("mFoo2 As Long", refactoredCode.Source);
            StringAssert.DoesNotContain("mFoo3 As Long", refactoredCode.Source);
            StringAssert.Contains(destinationDeclaration, refactoredCode.Destination);
            StringAssert.Contains($"{endpoints.DestinationModuleName()}.mFoo", refactoredCode.Source);
            StringAssert.Contains($"{endpoints.DestinationModuleName()}.mFoo2", refactoredCode.Source);
            StringAssert.Contains($"{endpoints.DestinationModuleName()}.mFoo", refactoredCode[callSiteModuleName]);
            StringAssert.Contains($"result = ({endpoints.DestinationModuleName()}.mFoo + arg1) * 2", refactoredCode[callSiteModuleName]);
            StringAssert.Contains($"NonQualifiedFoo = ({endpoints.DestinationModuleName()}.mFoo2 + arg1) * 2", refactoredCode[callSiteModuleName]);
        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MoveNonAggregateValueTypeFieldsList_HasReference()
        {
            var endpoints = MoveEndpoints.StdToStd;
            var memberToMove = ("mFoo", DeclarationType.Variable);
            var source =
$@"
Option Explicit

Public mFoo As Long, mFoo1 As Long, mFoo2 As Long, mFoo3 As Long

Public Function FooMath(arg1 As Long) As Long
    FooMath = arg1 * mFoo * mFoo2
End Function
";
            var callSiteModuleName = "CallSiteModule";
            var otherModuleReference =
    $@"
Option Explicit

Public Function MemberAccessFoo(arg1 As Long) As Long
    MemberAccessFoo = arg1 * {endpoints.SourceModuleName()}.mFoo * 2
End Function

Public Function WithMemberAccessFoo(arg1 As Long) As Long
    Dim result As Long
    With {endpoints.SourceModuleName()}
        result = (.mFoo + arg1) * 2
    End With
    WithMemberAccessFoo = result
End Function

Public Function NonQualifiedFoo(arg1 As Long) As Long
    NonQualifiedFoo = (mFoo2 + arg1) * 2
End Function
";
            MoveMemberModel ModelAdjustment(MoveMemberModel model)
            {
                model.MoveableMemberSetByName("mFoo1").IsSelected = true;
                model.MoveableMemberSetByName("mFoo2").IsSelected = true;
                model.MoveableMemberSetByName("mFoo3").IsSelected = true;
                model.ChangeDestination(endpoints.DestinationModuleName(), ComponentType.StandardModule);
                return model;
            }

            var refactoredCode = RefactorTargets(memberToMove, endpoints, ModelAdjustment,
                endpoints.ToSourceTuple(source),
                endpoints.ToDestinationTuple(string.Empty),
                (callSiteModuleName, otherModuleReference, ComponentType.StandardModule));

            Assert.IsTrue(Tokens.Public.OccursOnce(refactoredCode.Source));
            StringAssert.DoesNotContain("mFoo As Long", refactoredCode.Source);
            StringAssert.DoesNotContain("mFoo1 As Long", refactoredCode.Source);
            StringAssert.DoesNotContain("mFoo2 As Long", refactoredCode.Source);
            StringAssert.DoesNotContain("mFoo3 As Long", refactoredCode.Source);
            StringAssert.Contains("Public mFoo As Long", refactoredCode.Destination);
            StringAssert.Contains($"Public mFoo1 As Long", refactoredCode.Destination);
            StringAssert.Contains("Public mFoo2 As Long", refactoredCode.Destination);
            StringAssert.Contains($"Public mFoo3 As Long", refactoredCode.Destination);
            StringAssert.Contains($"{endpoints.DestinationModuleName()}.mFoo", refactoredCode.Source);
            StringAssert.Contains($"{endpoints.DestinationModuleName()}.mFoo2", refactoredCode.Source);
            StringAssert.Contains($"{endpoints.DestinationModuleName()}.mFoo", refactoredCode[callSiteModuleName]);
            StringAssert.Contains($"result = ({endpoints.DestinationModuleName()}.mFoo + arg1) * 2", refactoredCode[callSiteModuleName]);
            StringAssert.Contains($"NonQualifiedFoo = ({endpoints.DestinationModuleName()}.mFoo2 + arg1) * 2", refactoredCode[callSiteModuleName]);
        }

        [TestCase(MoveEndpoints.FormToStd, "Private")]
        [TestCase(MoveEndpoints.ClassToStd, "Private")]
        [TestCase(MoveEndpoints.StdToStd, "Public")]
        [TestCase(MoveEndpoints.StdToStd, "Private")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MoveNConstantsList_HasReferences(MoveEndpoints endpoints, string accessibility)
        {
            var memberToMove = ("mFoo", DeclarationType.Constant);
            var source =
$@"
Option Explicit

{accessibility} Const mFoo As Long = 0, mFoo1 As Long = 10, mFoo2 As Long = 20, mFoo3 As Long = 30

Public Function FooMath(arg1 As Long) As Long
    FooMath = arg1 * mFoo * mFoo2
End Function
";
            var callSiteModuleName = "CallSiteModule";
            var modules = new List<(string, string, ComponentType)>();
            if (endpoints.IsStdModuleSource() && accessibility.Equals(Tokens.Public))
            {
                var otherModuleReference =
    $@"
Option Explicit

Public Function MemberAccessFoo(arg1 As Long) As Long
    MemberAccessFoo = arg1 * {endpoints.SourceModuleName()}.mFoo * 2
End Function

Public Function WithMemberAccessFoo(arg1 As Long) As Long
    Dim result As Long
    With {endpoints.SourceModuleName()}
        result = (.mFoo + arg1) * 2
    End With
    WithMemberAccessFoo = result
End Function

Public Function NonQualifiedFoo(arg1 As Long) As Long
    NonQualifiedFoo = (mFoo2 + arg1) * 2
End Function
";
                modules.Add((callSiteModuleName, otherModuleReference, ComponentType.StandardModule));
            }

            modules.Add(endpoints.ToSourceTuple(source));
            modules.Add(endpoints.ToDestinationTuple(string.Empty));

            MoveMemberModel ModelAdjustment(MoveMemberModel model)
            {
                model.MoveableMemberSetByName("mFoo1").IsSelected = true;
                model.MoveableMemberSetByName("mFoo2").IsSelected = true;
                model.MoveableMemberSetByName("mFoo3").IsSelected = true;
                model.ChangeDestination(endpoints.DestinationModuleName(), ComponentType.StandardModule);
                return model;
            }

            var refactoredCode = RefactorTargets(memberToMove, endpoints, ModelAdjustment, modules.ToArray());

            StringAssert.DoesNotContain($"{accessibility} Const", refactoredCode.Source);
            StringAssert.DoesNotContain("mFoo As Long", refactoredCode.Source);
            StringAssert.DoesNotContain("mFoo1 As Long", refactoredCode.Source);
            StringAssert.DoesNotContain("mFoo2 As Long", refactoredCode.Source);
            StringAssert.DoesNotContain("mFoo3 As Long", refactoredCode.Source);
            StringAssert.Contains("Public Const mFoo As Long", refactoredCode.Destination);
            StringAssert.Contains($"{accessibility} Const mFoo1 As Long", refactoredCode.Destination);
            StringAssert.Contains("Public Const mFoo2 As Long", refactoredCode.Destination);
            StringAssert.Contains($"{accessibility} Const mFoo3 As Long", refactoredCode.Destination);
            StringAssert.Contains($"{endpoints.DestinationModuleName()}.mFoo", refactoredCode.Source);
            StringAssert.Contains($"{endpoints.DestinationModuleName()}.mFoo2", refactoredCode.Source);
            if (endpoints.IsStdModuleSource() && accessibility.Equals(Tokens.Public))
            {
                StringAssert.Contains($"{endpoints.DestinationModuleName()}.mFoo", refactoredCode[callSiteModuleName]);
                StringAssert.Contains($"result = ({endpoints.DestinationModuleName()}.mFoo + arg1) * 2", refactoredCode[callSiteModuleName]);
                StringAssert.Contains($"NonQualifiedFoo = ({endpoints.DestinationModuleName()}.mFoo2 + arg1) * 2", refactoredCode[callSiteModuleName]);
            }
        }

        [TestCase("PVT_VALUE", true)]
        [TestCase("85 + PVT_VALUE * 4", true)]
        [TestCase("PUB_VALUE + PVT_VALUE", true)]
        [TestCase("PUB_VALUE", false)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MoveConstantDeclarationReferencesOtherConstant(string expression, bool throwsException)
        {
            var endpoints = MoveEndpoints.StdToStd;
            var memberToMove = ("FIZZ", DeclarationType.Constant);
            var source =
$@"
Option Explicit

Private Const PVT_VALUE As Long = 75

Public Const PUB_VALUE As Long = 10

Public Const PUB_VALUE2 As Long = 20

Public Const FIZZ As Long = {expression}

Private Function Bizz(arg As Long) As Long
    Bizz = arg + PVT_VALUE
End Function
";
            if (throwsException)
            {
                ExecuteSingleTargetMoveThrowsExceptionTest(memberToMove, endpoints, source);
                return;
            }

            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source);

            StringAssert.Contains($"Public Const FIZZ As Long = {endpoints.SourceModuleName()}.{expression}", refactoredCode.Destination);
        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MoveConstantDeclarationReferencesOtherConstants()
        {
            var endpoints = MoveEndpoints.StdToStd;
            var memberToMove = ("FIZZ", DeclarationType.Constant);
            var source =
$@"
Option Explicit

Private Const PVT_VALUE As Long = 75

Public Const PUB_VALUE As Long = 10

Public Const PUB_VALUE2 As Long = 20

Public Const FIZZ As Long = PUB_VALUE * PUB_VALUE2

Private Function Bizz(arg As Long) As Long
    Bizz = arg + PVT_VALUE
End Function
";
            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source);
            StringAssert.Contains($"Public Const FIZZ As Long = {endpoints.SourceModuleName()}.PUB_VALUE * {endpoints.SourceModuleName()}.PUB_VALUE2", refactoredCode.Destination);
        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MoveConstantDeclarationAndAllSupport()
        {
            var endpoints = MoveEndpoints.StdToStd;
            var memberToMove = ("FIZZ", DeclarationType.Constant);
            var source =
$@"
Option Explicit

Private Const PVT_VALUE As Long = 75

Public Const FIZZ As Long = PVT_VALUE

Public Type TestType
    TheValue As Long
End Type

Private Function Bizz(arg As Long) As Long
    Bizz = arg + PVT_VALUE
End Function
";

            MoveMemberModel ModelAdjustment(MoveMemberModel model)
            {
                model.MoveableMemberSetByName("PVT_VALUE").IsSelected = true;
                model.MoveableMemberSetByName("Bizz").IsSelected = true;
                model.ChangeDestination(endpoints.DestinationModuleName(), ComponentType.StandardModule);
                return model;
            }

            var refactoredCode = RefactorTargets(memberToMove, endpoints, source, string.Empty, ModelAdjustment);
            StringAssert.Contains("Public Type TestType", refactoredCode.Source);

            StringAssert.Contains("Private Const PVT_VALUE As Long = 75", refactoredCode.Destination);
            StringAssert.Contains("Public Const FIZZ As Long = PVT_VALUE", refactoredCode.Destination);
            StringAssert.Contains("Private Function Bizz(arg As Long) As Long", refactoredCode.Destination);
        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MovePrivateConstantDeclarationWithSourceReferences()
        {
            var endpoints = MoveEndpoints.StdToStd;
            var memberToMove = ("PVT_VALUE", DeclarationType.Constant);
            var source =
$@"
Option Explicit

Private Const PVT_VALUE As Long = 75

Public Const FIZZ As Long = PVT_VALUE

Public Type TestType
    TheValue As Long
End Type

Private Function Bizz(arg As Long) As Long
    Bizz = arg + PVT_VALUE
End Function
";

            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source);

            StringAssert.Contains($"Public Const FIZZ As Long = {endpoints.DestinationModuleName()}.PVT_VALUE", refactoredCode.Source);
            StringAssert.Contains("Public Type TestType", refactoredCode.Source);
            StringAssert.Contains($"Bizz = arg + {endpoints.DestinationModuleName()}.PVT_VALUE", refactoredCode.Source);

            StringAssert.Contains("Public Const PVT_VALUE As Long = 75", refactoredCode.Destination);
            StringAssert.DoesNotContain("Public Const FIZZ As Long = PVT_VALUE", refactoredCode.Destination);
            StringAssert.DoesNotContain("Private Function Bizz(arg As Long) As Long", refactoredCode.Destination);
        }

        [TestCase(MoveEndpoints.FormToStd, "Public")]
        [TestCase(MoveEndpoints.FormToStd, "Private")]
        [TestCase(MoveEndpoints.ClassToStd, "Public")]
        [TestCase(MoveEndpoints.ClassToStd, "Private")]
        [TestCase(MoveEndpoints.StdToStd, "Private")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MoveObjectField_MoveToStdModuleStrategyNotApplicable(MoveEndpoints endpoints, string accessibility)
        {
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

            var state = CreateAndParse(endpoints, source, string.Empty);
            using (state)
            {
                var resolver = new MoveMemberTestsResolver(state);
                var strategyFactory = resolver.Resolve<IMoveMemberStrategyFactory>();
                var model = MoveMemberTestsResolver.CreateRefactoringModel("mObj", DeclarationType.Variable, state);

                var strategy = strategyFactory.Create(model.MoveEndpoints);
                Assert.False(strategy.IsApplicable(model));
            }
        }


        [TestCase("Public mObj As ObjectClass")]
        [TestCase("Public mObj As New ObjectClass")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void FunctionReferencesPublicObjectField(string declaration)
        {
            var memberToMove = ("mObj", DeclarationType.Variable);
            var endpoints = MoveEndpoints.StdToStd;
            var source =
$@"
Option Explicit

{declaration}

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
            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source);

            StringAssert.DoesNotContain($"{declaration}", refactoredCode.Source);
            StringAssert.Contains($"if {endpoints.DestinationModuleName()}.mObj is Nothing Then", refactoredCode.Source);
            StringAssert.Contains($"Set {endpoints.DestinationModuleName()}.mObj = new ObjectClass", refactoredCode.Source);
            StringAssert.Contains($"FooMath = arg1 * {endpoints.DestinationModuleName()}.mObj.Value", refactoredCode.Source);

            StringAssert.Contains($"{declaration}", refactoredCode.Destination);
        }

        [Test]
        [Category("MoveMember")]
        public void StdToStdPublicArrayMoveFieldHasReferences()
        {
            var endpoints = MoveEndpoints.StdToStd;
            var memberToMove = ("mArray", DeclarationType.Variable);
            var source =
$@"
Option Explicit

Public mArray(5) As Long

Public Function FooMath(arg1 As Long) As Long
    FooMath = arg1 * mArray(2)
End Function
";
            var callSiteModuleName = "CallSiteModule";

            var otherModuleReference =
    $@"
Option Explicit

Public Function MemberAccessFoo(arg1 As Long) As Long
    MemberAccessFoo = arg1 + {endpoints.SourceModuleName()}.mArray(3) * 2
End Function

Public Function WithMemberAccessFoo(arg1 As Long) As Long
    Dim result As Long
    With {endpoints.SourceModuleName()}
        result = (.mArray(2) + arg1) * 2
    End With
    WithMemberAccessFoo = result
End Function

Public Function NonQualifiedFoo(arg1 As Long) As Long
    NonQualifiedFoo = (mArray(1) + arg1) * 2
End Function
";
            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints,
                endpoints.ToSourceTuple(source),
                endpoints.ToDestinationTuple(string.Empty),
                (callSiteModuleName, otherModuleReference, ComponentType.StandardModule));

            var destinationDeclaration = "Public mArray(5) As Long";

            StringAssert.DoesNotContain("mArray(5) As Long", refactoredCode.Source);
            StringAssert.Contains(destinationDeclaration, refactoredCode.Destination);
            StringAssert.Contains($"{endpoints.DestinationModuleName()}.mArray(2)", refactoredCode.Source);
            StringAssert.Contains($"{endpoints.DestinationModuleName()}.mArray(3)", refactoredCode[callSiteModuleName]);
            StringAssert.Contains($"result = ({endpoints.DestinationModuleName()}.mArray(2) + arg1) * 2", refactoredCode[callSiteModuleName]);
            StringAssert.Contains($"NonQualifiedFoo = ({endpoints.DestinationModuleName()}.mArray(1) + arg1) * 2", refactoredCode[callSiteModuleName]);
        }

        [Test]
        [Category(nameof(ConflictDetectionSession))]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void CorrectsFieldNameCollisionInDestination()
        {
            var endpoints = MoveEndpoints.StdToStd;
            var memberToMove = ("Goo", DeclarationType.PropertyLet);
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
            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source, destination);

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
                StringAssert.Contains(line, refactoredCode.Destination);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void SetsNewFieldNameAtExternalReferences()
        {
            var destinationModuleName = MoveEndpoints.StdToStd.DestinationModuleName();
            var sourceModuleName = MoveEndpoints.StdToStd.SourceModuleName();
            var memberToMove = ("mfoo", DeclarationType.Variable);
            var endpoints = MoveEndpoints.StdToStd;
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
    mBar = {sourceModuleName}.mfoo + arg1
End Sub

Public Sub WithMemberAccess(arg2 As Long)
    With {sourceModuleName}
        mBar = .mfoo + arg2
    End With
End Sub

Public Sub NonQualified(arg3 As Long)
    mBar = mfoo + arg3
End Sub
";
            var expectedCallSiteCode =
    $@"
Option Explicit

Private mBar As Long

Public Sub MemberAccess(arg1 As Long)
    mBar = {destinationModuleName}.mfoo1 + arg1
End Sub

Public Sub WithMemberAccess(arg2 As Long)
    With {sourceModuleName}
        mBar = {destinationModuleName}.mfoo1 + arg2
    End With
End Sub

Public Sub NonQualified(arg3 As Long)
    mBar = {destinationModuleName}.mfoo1 + arg3
End Sub
";
            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, endpoints.SourceModuleName(), 
                endpoints.ToSourceTuple(source),
                endpoints.ToDestinationTuple(destination),
                (callSiteModuleName, callSiteCode, ComponentType.StandardModule));

            var expectedLines = expectedCallSiteCode.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var line in expectedLines)
            {
                StringAssert.Contains(line, refactoredCode[callSiteModuleName]);
            }
        }

        [Test]
        [Category(nameof(ConflictDetectionSession))]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void CorrectsConstantNameCollisionInDestination()
        {
            var memberToMove = ("Goo", DeclarationType.PropertyLet);
            var endpoints = MoveEndpoints.StdToStd;
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
            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, endpoints.SourceModuleName(), source, destination);


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
                StringAssert.Contains(line, refactoredCode.Destination);
            }
        }

        [TestCase("mfoo", false)]
        [TestCase("BigFoo", true)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void SetsNewConstantNameAtExternalReferences(string selectedConstant, bool useModuleQualification /*expectedModuleQualification*/)
        {
            var memberToMove = (selectedConstant, DeclarationType.Constant);
            var endpoints = MoveEndpoints.StdToStd;
            var sourceModuleName = MoveEndpoints.StdToStd.SourceModuleName();
            var destinationModuleName = MoveEndpoints.StdToStd.DestinationModuleName();
            var source =
$@"
Option Explicit

Public Const mfoo As Long = 5
Public Const BigFoo As Long = mfoo
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
    mBar = {sourceModuleName}.BigFoo + arg1
End Sub

Public Sub WithMemberAccess(arg2 As Long)
    With {sourceModuleName}
        mBar = .BigFoo + arg2
    End With
End Sub

Public Sub NonQualified(arg3 As Long)
    mBar = BigFoo + arg3
End Sub
";
            var memberAccessQualification = useModuleQualification
                ? $"{destinationModuleName}."
                : $"{sourceModuleName}.";

            var withMemberAccessQualification = useModuleQualification
                ? $"{destinationModuleName}."
                : ".";

            var nonQualifiedAccess = useModuleQualification
                ? $"{destinationModuleName}."
                : string.Empty;

            var expectedCallSiteCode =
    $@"
Option Explicit

Private mBar As Long

Public Sub MemberAccess(arg1 As Long)
    mBar = {memberAccessQualification}BigFoo + arg1
End Sub

Public Sub WithMemberAccess(arg2 As Long)
    With {sourceModuleName}
        mBar = {withMemberAccessQualification}BigFoo + arg2
    End With
End Sub

Public Sub NonQualified(arg3 As Long)
    mBar = {nonQualifiedAccess}BigFoo + arg3
End Sub
";
            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, 
                    endpoints.ToSourceTuple(source),
                    endpoints.ToDestinationTuple(string.Empty),
                    (callSiteModuleName, callSiteCode, ComponentType.StandardModule));

            var expectedLines = expectedCallSiteCode.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var line in expectedLines)
            {
                StringAssert.Contains(line, refactoredCode[callSiteModuleName]);
            }
        }
    }
}
