using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System;
using Support = RubberduckTests.Refactoring.MoveMember.MoveMemberTestSupport;

namespace RubberduckTests.Refactoring.MoveMember
{
    /*
     * Every test in this class should have an 
     * equivalent test in the SingleProcedureToStdModuleTests class 
     */
    [TestFixture]
    public class MoveFunctionsToStdModuleTests : MoveMemberTestsBase
    {
        private const string ThisStrategy = nameof(MoveMemberToStdModule);
        private const DeclarationType ThisDeclarationType = DeclarationType.Function;

        [TestCase("Public", MoveEndpoints.StdToClass)]
        [TestCase("Private", MoveEndpoints.StdToClass)]
        [TestCase("Public", MoveEndpoints.ClassToClass)]
        [TestCase("Private", MoveEndpoints.ClassToClass)]
        [TestCase("Public", MoveEndpoints.FormToClass)]
        [TestCase("Private", MoveEndpoints.FormToClass)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void SimpleMoveToClassModule_NoStrategy(string accessibility, MoveEndpoints endpoints)
        {
            var memberToMove = "Foo";
            var source =
$@"
Option Explicit

{accessibility} Function Foo(arg1 As Long) As Long
    Foo = 10 * arg1
End Function
";

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType));

            var refactoredCode = RefactoredCode_UserSetsDestinationModuleName(moveDefinition, source);

            StringAssert.AreEqualIgnoringCase(null, refactoredCode.StrategyName);
        }

        [TestCase(MoveEndpoints.StdToStd)]
        [TestCase(MoveEndpoints.ClassToStd)]
        [TestCase(MoveEndpoints.FormToStd)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void PreviewMovedContent(MoveEndpoints endpoints)
        {
            var memberToMove = ("Foo", ThisDeclarationType);
            var source =
$@"
Option Explicit

Function Foo(arg1 As Long) As Long
    Const localConst As Long = 5
    Dim local As Long
    local = 6
    Foo = localConst + localVar + arg1
End Function
";

            var moveDefinition = new TestMoveDefinition(endpoints, memberToMove);
            var preview = RetrievePreviewAfterUserInput(moveDefinition, source, memberToMove);

            StringAssert.Contains("Option Explicit", preview);
            Assert.IsTrue(Support.OccursOnce("Public Function Foo(", preview));
        }

        [TestCase(MoveEndpoints.StdToStd, "Public", ThisStrategy)]
        [TestCase(MoveEndpoints.StdToStd, "Private", ThisStrategy)]
        [TestCase(MoveEndpoints.ClassToStd, "Public", ThisStrategy)]
        [TestCase(MoveEndpoints.ClassToStd, "Private", ThisStrategy)]
        [TestCase(MoveEndpoints.FormToStd, "Public", ThisStrategy)]
        [TestCase(MoveEndpoints.FormToStd, "Private", ThisStrategy)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MovedFunctionReferencesExclusiveSupportConstant(MoveEndpoints endpoints, string exclusiveFuncAccessibility, string expectedStrategy)
        {
            var memberToMove = "CalculateVolumeFromDiameter";
            var pi = "Pi";
            var source =
$@"
Option Explicit

Private Const {pi} As Single = 3.14

Public Function CalculateVolumeFromDiameter(ByVal diameter As Single, ByVal height As Single) As Single
    CalculateVolumeFromDiameter = CalculateVolume(diameter / 2, height)
End Function

{exclusiveFuncAccessibility} Function CalculateVolume(ByVal radius As Single, ByVal height As Single) As Single
    CalculateVolume = height * {pi} * radius ^ 2
End Function
";

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType));

            var refactoredCode = RefactoredCode_UserSetsDestinationModuleName(moveDefinition, source);

            StringAssert.AreEqualIgnoringCase(expectedStrategy, refactoredCode.StrategyName);

            if (expectedStrategy is null)
            {
                return;
            }

            StringAssert.DoesNotContain("Public Function CalculateVolumeFromDiameter(", refactoredCode.Source);
            StringAssert.DoesNotContain($"{exclusiveFuncAccessibility} Function CalculateVolume(", refactoredCode.Source);
            StringAssert.DoesNotContain($"{pi} As Single", refactoredCode.Source);

            StringAssert.Contains($"Private Const {pi} As Single", refactoredCode.Destination);
            StringAssert.Contains("Public Function CalculateVolumeFromDiameter(", refactoredCode.Destination);
            StringAssert.Contains("CalculateVolumeFromDiameter = CalculateVolume(", refactoredCode.Destination);
            StringAssert.Contains($"{exclusiveFuncAccessibility} Function CalculateVolume(", refactoredCode.Destination);
        }

        [TestCase(MoveEndpoints.StdToStd, ThisStrategy)]
        [TestCase(MoveEndpoints.ClassToStd, null)]
        [TestCase(MoveEndpoints.FormToStd, null)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MovedFunctionReferencesNonExclusivePublicSupportConstant(MoveEndpoints endpoints, string expectedStrategy)
        {
            var memberToMove = "CalculateVolumeFromDiameter";
            var pi = "Pi";
            var source =
$@"
Option Explicit

Public Const {pi} As Single = 3.14

Public Function CalculateVolumeFromDiameter(ByVal diameter As Single, ByVal height As Single) As Single
    CalculateVolumeFromDiameter = height * {pi} * (diameter / 2) ^ 2
End Function

Public Function CalculateCircumferenceFromDiameter(ByVal diameter As Single) As Single
    CalculateCircumferenceFromDiameter = diameter * {pi}
End Function";

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType));

            var refactoredCode = RefactoredCode_UserSetsDestinationModuleName(moveDefinition, source);

            StringAssert.AreEqualIgnoringCase(expectedStrategy, refactoredCode.StrategyName);

            if (expectedStrategy is null)
            {
                return;
            }

            StringAssert.DoesNotContain("Public Function CalculateVolumeFromDiameter(", refactoredCode.Source);
            StringAssert.Contains($"Public Function CalculateCircumferenceFromDiameter(", refactoredCode.Source);

            StringAssert.Contains($"Public Function CalculateVolumeFromDiameter(", refactoredCode.Destination);
            StringAssert.Contains($"height * {moveDefinition.SourceModuleName}.{pi}", refactoredCode.Destination);
        }

        [TestCase(MoveEndpoints.StdToStd, null)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MovedFunctionReferencesNonExclusivePrivateSupportConstant(MoveEndpoints endpoints, string expectedStrategy)
        {
            var memberToMove = "CalculateVolumeFromDiameter";
            var pi = "Pi";
            var source =
$@"
Option Explicit

Private Const {pi} As Single = 3.14

Public Function CalculateVolumeFromDiameter(ByVal diameter As Single, ByVal height As Single) As Single
    CalculateVolumeFromDiameter = height * {pi} * (diameter / 2) ^ 2
End Function

Public Function CalculateCircumferenceFromDiameter(ByVal diameter As Single) As Single
    CalculateCircumferenceFromDiameter = diameter * {pi}
End Function";

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType));

            var refactoredCode = RefactoredCode_UserSetsDestinationModuleName(moveDefinition, source);

            StringAssert.AreEqualIgnoringCase(expectedStrategy, refactoredCode.StrategyName);
        }

        [TestCase("Public", MoveEndpoints.StdToStd, ThisStrategy)]
        [TestCase("Private", MoveEndpoints.StdToStd, ThisStrategy)]
        [TestCase("Public", MoveEndpoints.ClassToStd, null)]
        [TestCase("Private", MoveEndpoints.ClassToStd, null)]
        [TestCase("Public", MoveEndpoints.FormToStd, null)]
        [TestCase("Private", MoveEndpoints.FormToStd, null)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ReferencesNonExclusivePublicField(string accessibility, MoveEndpoints endpoints, string expectedStrategy)
        {
            var memberToMove = "Foo";
            var source = $@"
Option Explicit

Public bar As Long

{accessibility} Function Foo(arg1 As Long) As Long
    Foo = bar + arg1
End Function

Public Sub Goo(arg1 As Long)
    bar = bar + arg1
End Sub
";

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType));

            var refactoredCode = RefactoredCode_UserSetsDestinationModuleName(moveDefinition, source);

            StringAssert.AreEqualIgnoringCase(expectedStrategy, refactoredCode.StrategyName);

            if (expectedStrategy is null)
            {
                return;
            }

            StringAssert.DoesNotContain($"{accessibility} Function Foo(", refactoredCode.Source);
            StringAssert.Contains($"Public Sub Goo(", refactoredCode.Source);
            StringAssert.Contains($"Public bar As Long", refactoredCode.Source);

            StringAssert.Contains($"Public Function Foo(", refactoredCode.Destination);
            StringAssert.Contains($"Foo = {moveDefinition.SourceModuleName}.bar + arg1", refactoredCode.Destination);
            StringAssert.DoesNotContain($"Public bar As Long", refactoredCode.Destination);
        }

        [TestCase(MoveEndpoints.StdToStd, "Public", ThisStrategy)]
        [TestCase(MoveEndpoints.StdToStd, "Private", null)]
        [TestCase(MoveEndpoints.ClassToStd, "Public", null)]
        [TestCase(MoveEndpoints.ClassToStd, "Private", null)]
        [TestCase(MoveEndpoints.FormToStd, "Public", null)]
        [TestCase(MoveEndpoints.FormToStd, "Private", null)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ReferencesNonExclusiveMember(MoveEndpoints endpoints, string exclusiveFuncAccessibility, string expectedStrategy)
        {
            var memberToMove = "CalculateVolumeFromDiameter";
            var source =
$@"
Option Explicit

Public Function CalculateVolumeFromDiameter(ByVal diameter As Single, ByVal height As Single) As Single
    CalculateVolumeFromDiameter = CalculateVolume(diameter / 2, height)
End Function

{exclusiveFuncAccessibility} Function CalculateVolume(ByVal radius As Single, ByVal height As Single) As Single
    CalculateVolume = height * 3.14 * radius ^ 2
End Function

Public Function CalculateVolumeFromCyliderCircumference(ByVal circumference As Single, height As Single) As Single
    CalculateVolumeFromCyliderCircumference = CalculateVolume((circumference / Pi) / 2, height)
End Function
";

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType));

            var refactoredCode = RefactoredCode_UserSetsDestinationModuleName(moveDefinition, source);

            StringAssert.AreEqualIgnoringCase(expectedStrategy, refactoredCode.StrategyName);

            if (expectedStrategy is null)
            {
                return;
            }

            StringAssert.DoesNotContain("Public Function CalculateVolumeFromDiameter(", refactoredCode.Source);
            StringAssert.Contains($"{exclusiveFuncAccessibility} Function CalculateVolume(", refactoredCode.Source);

            StringAssert.Contains($"Public Function CalculateVolumeFromDiameter(", refactoredCode.Destination);
            StringAssert.Contains($"CalculateVolumeFromDiameter = {moveDefinition.SourceModuleName}.CalculateVolume(", refactoredCode.Destination);
            StringAssert.DoesNotContain($"{exclusiveFuncAccessibility} Function CalculateVolume(", refactoredCode.Destination);
        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ExternallyReferencedMemberStdToStd()
        {
            var memberToMove = "CalculateVolumeFromDiameter";
            var source =
$@"
Option Explicit

Public Function CalculateVolumeFromDiameter(ByVal diameter As Single, ByVal height As Single) As Single
    CalculateVolumeFromDiameter = height * 3.14 * (diameter / 2) ^ 2 
End Function
";


            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, (memberToMove, ThisDeclarationType));

            var externalReferences =
$@"
Option Explicit

Private mFoo As Single

Public Sub MemberAccess()
    mFoo = {moveDefinition.SourceModuleName}.CalculateVolumeFromDiameter(7.5, 4.2)
End Sub

Public Sub WithMemberAccess()
    With {moveDefinition.SourceModuleName}
        mFoo = .CalculateVolumeFromDiameter(8.5, 4.2)
    End With
End Sub

Public Sub NonQualifiedAccess()
    mFoo = CalculateVolumeFromDiameter(9.5, 4.2)
End Sub
";
            var referencingModuleName = "Module3";
            moveDefinition.Add(new ModuleDefinition(referencingModuleName, ComponentType.StandardModule, externalReferences));

            var refactoredCode = RefactoredCode_UserSetsDestinationModuleName(moveDefinition, source);

            StringAssert.AreEqualIgnoringCase(ThisStrategy, refactoredCode.StrategyName);

            StringAssert.DoesNotContain("Public Function CalculateVolumeFromDiameter(", refactoredCode.Source);
            StringAssert.Contains($"Public Function CalculateVolumeFromDiameter(", refactoredCode.Destination);

            var module3Content = refactoredCode[referencingModuleName];

            StringAssert.Contains($"{moveDefinition.DestinationModuleName}.CalculateVolumeFromDiameter(7.5, 4.2)", module3Content);
            StringAssert.Contains($"With {moveDefinition.SourceModuleName}", module3Content);
            StringAssert.Contains($"{moveDefinition.DestinationModuleName}.CalculateVolumeFromDiameter(8.5, 4.2)", module3Content);
            StringAssert.Contains($"{moveDefinition.DestinationModuleName}.CalculateVolumeFromDiameter(9.5, 4.2)", module3Content);
        }

        [TestCase(MoveEndpoints.ClassToStd)]
        [TestCase(MoveEndpoints.FormToStd)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ExternallyReferencedMemberClassSource(MoveEndpoints endpoints)
        {
            var memberToMove = "CalculateVolumeFromDiameter";
            var source =
$@"
Option Explicit

Public Function CalculateVolumeFromDiameter(ByVal diameter As Single, ByVal height As Single) As Single
    CalculateVolumeFromDiameter = height * 3.14 * (diameter / 2) ^ 2 
End Function
";

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType));

            var instanceIdentifier = "mClassOrUserForm";
            var externalReferences =
$@"
Option Explicit

Private mFoo As Single

{Support.ClassInstantiationBoilerPlate(instanceIdentifier, moveDefinition.SourceModuleName)}

Public Sub MemberAccess()
    mFoo = {instanceIdentifier}.CalculateVolumeFromDiameter(7.5, 4.2)
End Sub

Public Sub WithMemberAccess()
    With {instanceIdentifier}
        mFoo = .CalculateVolumeFromDiameter(8.5, 4.2)
    End With
End Sub
";
            var referencingModuleName = "Module3";
            moveDefinition.Add(new ModuleDefinition(referencingModuleName, ComponentType.StandardModule, externalReferences));

            var refactoredCode = RefactoredCode_UserSetsDestinationModuleName(moveDefinition, source);

            StringAssert.AreEqualIgnoringCase(null, refactoredCode.StrategyName);
        }

        [TestCase(MoveEndpoints.StdToStd, "Public")]
        [TestCase(MoveEndpoints.StdToStd, "Private")]
        [TestCase(MoveEndpoints.ClassToStd, "Public")]
        [TestCase(MoveEndpoints.ClassToStd, "Private")]
        [TestCase(MoveEndpoints.FormToStd, "Public")]
        [TestCase(MoveEndpoints.FormToStd, "Private")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void SelfContainedMember(MoveEndpoints endpoints, string targetAccessibility)
        {
            var memberToMove = "TryRenameFile";
            var source =
$@"
Option Explicit

Private mfileName As String

Public Sub ChangeLogFile(fileName As String)
    If TryRenameFile(mfileName, fileName) Then
        mfileName = fileName
    EndIf
End Sub

{targetAccessibility} Function TryRenameFile(oldName As String, newName As String) As Boolean
    Name oldName As newName
    TryRenameFile = True
End Function
";

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType));

            var refactoredCode = RefactoredCode_UserSetsDestinationModuleName(moveDefinition, source);

            StringAssert.AreEqualIgnoringCase(ThisStrategy, refactoredCode.StrategyName);

            var sourceRefactored = refactoredCode.Source;
            StringAssert.Contains($"{moveDefinition.DestinationModuleName}.TryRenameFile(mfileName, fileName)", sourceRefactored);
            StringAssert.DoesNotContain($"{targetAccessibility} Function TryRenameFile", sourceRefactored);

            StringAssert.Contains("Public Function TryRenameFile", refactoredCode.Destination);
        }

        [TestCase(MoveEndpoints.StdToStd)]
        [TestCase(MoveEndpoints.ClassToStd)]
        [TestCase(MoveEndpoints.FormToStd)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void InternallyReferencedFunctionToStdModule(MoveEndpoints endpoints)
        {
            var memberToMove = "CalculateVolume";

            var source =
$@"
Option Explicit

Private Const Pi As Single = 3.14

Public Function CalculateCylinderVolumeFromDiameter(diameter As Single, height As Single) As Single
    CalculateCylinderVolumeFromDiameter = CalculateVolume(Pi * (diameter / 2) ^ 2, height)
End Function

Public Function CalculateCylinderVolumeFromRadius(radius As Single, height As Single) As Single
    CalculateCylinderVolumeFromRadius = CalculateVolume(Pi * (radius) ^ 2, height)
End Function

Private Function CalculateVolume(area As Single, height As Single) As Single
    CalculateVolume = area * height
End Function
";

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType));

            var refactoredCode = RefactoredCode_UserSetsDestinationModuleName(moveDefinition, source);

            StringAssert.AreEqualIgnoringCase(ThisStrategy, refactoredCode.StrategyName);
            StringAssert.Contains($"{moveDefinition.DestinationModuleName}.CalculateVolume(", refactoredCode.Source);
            StringAssert.DoesNotContain("Private Function CalculateVolume", refactoredCode.Source);

            StringAssert.Contains($"CalculateVolume = area * height", refactoredCode.Destination);
            StringAssert.Contains($"Public Function CalculateVolume(ByRef area", refactoredCode.Destination);
        }

        [TestCase(MoveEndpoints.StdToStd, "Public")]
        [TestCase(MoveEndpoints.StdToStd, "Private")]
        [TestCase(MoveEndpoints.ClassToStd, "Public")]
        [TestCase(MoveEndpoints.ClassToStd, "Private")]
        [TestCase(MoveEndpoints.FormToStd, "Public")]
        [TestCase(MoveEndpoints.FormToStd, "Private")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ExclusiveCallChain(MoveEndpoints endpoints, string accessibility)
        {
            var memberToMove = "Foo";
            var source =
$@"
Option Explicit

Private mfoo5 As Long, mfoo As Long, mfoo2 As Long, mfoo3 As Long, mfoo4 As Long

{accessibility} Function Foo(arg1 As Long) As Long
    mfoo = Bar(arg1)
    Foo = mfoo
End Function

Public Sub Goo(arg1 As Long)
    mfoo5 = AddSeven(arg1)
End Sub

Private Function AddSix(arg1 As Long) As Long
    AddSix = arg1 + 6
End Function

Private Function AddSeven(arg1 As Long) As Long
    AddSeven = arg1 + 7
End Function

Private Function Bar(arg1 As Long) As Long
    mfoo2 = arg1
    Bar = Barn(mfoo2)
End Function

Private Function Barn(arg1 As Long) As Long
    mfoo3 = arg1
    Barn = Bark(mfoo3)
End Function

Private Function Bark(arg1 As Long) As Long
    mfoo4 = AddSix(arg1)
    Bark = mfoo4
End Function
";

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType));

            var refactoredCode = RefactoredCode_UserSetsDestinationModuleName(moveDefinition, source);

            StringAssert.AreEqualIgnoringCase(ThisStrategy, refactoredCode.StrategyName);

            var sourceRefactored = refactoredCode.Source;
            StringAssert.DoesNotContain("Private Function Bar", sourceRefactored);
            StringAssert.DoesNotContain("Private Function Barn", sourceRefactored);
            StringAssert.DoesNotContain("Private Function Bark", sourceRefactored);
            StringAssert.DoesNotContain("Private mfoo As Long, mfoo2 As Long, mfoo3 As Long, mfoo4 As Long", sourceRefactored);
            StringAssert.DoesNotContain("Private Function AddSix(", sourceRefactored);
            StringAssert.Contains("Private Function AddSeven(", sourceRefactored);
            StringAssert.Contains("Private mfoo5 As Long", sourceRefactored);

            var destinationRefactored = refactoredCode.Destination;
            StringAssert.Contains("Private Function Bar", destinationRefactored);
            StringAssert.Contains("Private Function Barn", destinationRefactored);
            StringAssert.Contains("Private Function Bark", destinationRefactored);
            StringAssert.Contains(" mfoo As Long", destinationRefactored);
            StringAssert.Contains(" mfoo2 As Long", destinationRefactored);
            StringAssert.Contains("Private Function AddSix(", destinationRefactored);
        }

        [TestCase(MoveEndpoints.StdToStd, "Public")]
        [TestCase(MoveEndpoints.StdToStd, "Private")]
        [TestCase(MoveEndpoints.ClassToStd, "Public")]
        [TestCase(MoveEndpoints.ClassToStd, "Private")]
        [TestCase(MoveEndpoints.FormToStd, "Public")]
        [TestCase(MoveEndpoints.FormToStd, "Private")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ExclusiveCallChainNonExclusiveField_NoStrategy(MoveEndpoints endpoints, string accessibility)
        {
            var memberToMove = "Foo";
            var mfoo5 = "mfoo5";
            var source =
$@"
Option Explicit

Private {mfoo5} As Long, mfoo As Long, mfoo2 As Long, mfoo3 As Long, mfoo4 As Long

{accessibility} Function Foo(arg1 As Long) As Long
    mfoo = Bar(arg1)
    Foo = mfoo
End Function

Public Sub Goo(arg1 As Long)
    {mfoo5} = AddSeven(arg1)
End Sub

Private Function AddSix(arg1 As Long) As Long
    AddSix = arg1 + 6
End Function

Private Function AddSeven(arg1 As Long) As Long
    AddSeven = arg1 + 7
End Function

Private Function Bar(arg1 As Long) As Long
    mfoo2 = arg1
    Bar = Barn(mfoo2)
End Function

Private Function Barn(arg1 As Long) As Long
    mfoo3 = arg1
    Barn = Bark(mfoo3)
End Function

Private Function Bark(arg1 As Long) As Long
    mfoo4 = AddSix(arg1)
    Bark = mfoo4 + {mfoo5}
End Function
";

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType));

            var refactoredCode = RefactoredCode_UserSetsDestinationModuleName(moveDefinition, source);

            StringAssert.AreEqualIgnoringCase(null, refactoredCode.StrategyName);
        }

        [TestCase(MoveEndpoints.StdToStd, "Public")]
        [TestCase(MoveEndpoints.StdToStd, "Private")]
        [TestCase(MoveEndpoints.ClassToStd, "Public")]
        [TestCase(MoveEndpoints.ClassToStd, "Private")]
        [TestCase(MoveEndpoints.FormToStd, "Public")]
        [TestCase(MoveEndpoints.FormToStd, "Private")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void NonExclusiveCallChainPrivateMember_NoStrategy(MoveEndpoints endpoints, string accessibility)
        {
            var memberToMove = "Foo";
            var AddSix = "AddSix"; //NonExclusive callchain member
            var source =
$@"
Option Explicit

Private mfoo5 As Long, mfoo As Long, mfoo2 As Long, mfoo3 As Long, mfoo4 As Long

{accessibility} Function Foo(arg1 As Long)
    mfoo = Bar(arg1)
    Foo = mfoo
End Function

Public Sub Goo(arg1 As Long)
    mfoo5 = {AddSix}(arg1)
End Sub

Private Function AddSix(arg1 As Long) As Long
    AddSix = arg1 + 6
End Function

Private Function Bar(arg1 As Long) As Long
    mfoo2 = arg1
    Bar = Barn(mfoo2)
End Function

Private Function Barn(arg1 As Long) As Long
    mfoo3 = arg1
    Barn = Bark(mfoo3)
End Function

Private Function Bark(arg1 As Long) As Long
    mfoo4 = {AddSix}(arg1)
    Bark = mfoo4
End Function
";

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType));

            var refactoredCode = RefactoredCode_UserSetsDestinationModuleName(moveDefinition, source);

            StringAssert.AreEqualIgnoringCase(null, refactoredCode.StrategyName);
        }

        [TestCase(MoveEndpoints.StdToStd, "Public", ThisStrategy)]
        [TestCase(MoveEndpoints.ClassToStd, "Public", null)]
        [TestCase(MoveEndpoints.FormToStd, "Public", null)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void NonExclusiveCallChainPublicMember(MoveEndpoints endpoints, string accessibility, string expectedStrategy)
        {
            var memberToMove = "Foo";
            var AddSix = "AddSix"; //NonExclusive callchain member
            var source =
$@"
Option Explicit

Private mfoo5 As Long, mfoo As Long, mfoo2 As Long, mfoo3 As Long, mfoo4 As Long

Public Function Foo(arg1 As Long) As Long
    mfoo = Bar(arg1)
    Foo = mfoo
End Function

Public Sub Goo(arg1 As Long)
    mfoo5 = {AddSix}(arg1)
End Sub

Public Function AddSix(arg1 As Long) As Long
    AddSix = arg1 + 6
End Function

Private Function Bar(arg1 As Long) As Long
    mfoo2 = arg1
    Bar = Barn(mfoo2)
End Function

Private Function Barn(arg1 As Long) As Long
    mfoo3 = arg1
    Barn = Bark(mfoo3)
End Function

Private Function Bark(arg1 As Long) As Long
    mfoo4 = {AddSix}(arg1)
    Bark = mfoo4
End Function
";

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType));

            var refactoredCode = RefactoredCode_UserSetsDestinationModuleName(moveDefinition, source);

            StringAssert.AreEqualIgnoringCase(expectedStrategy, refactoredCode.StrategyName);

            if (expectedStrategy is null)
            {
                return;
            }

            var sourceRefactored = refactoredCode.Source;
            StringAssert.DoesNotContain("Private Function Bar", sourceRefactored);
            StringAssert.DoesNotContain("Private Function Barn", sourceRefactored);
            StringAssert.DoesNotContain("Private Function Bark", sourceRefactored);
            StringAssert.DoesNotContain("Private mfoo As Long, mfoo2 As Long, mfoo3 As Long, mfoo4 As Long", sourceRefactored);
            StringAssert.Contains("Public Function AddSix(", sourceRefactored);
            StringAssert.Contains("Private mfoo5 As Long", sourceRefactored);

            var destinationRefactored = refactoredCode.Destination;
            StringAssert.Contains("Private Function Bar", destinationRefactored);
            StringAssert.Contains("Private Function Barn", destinationRefactored);
            StringAssert.Contains("Private Function Bark", destinationRefactored);
            StringAssert.Contains(" mfoo As Long", destinationRefactored);
            StringAssert.Contains(" mfoo2 As Long", destinationRefactored);
            StringAssert.DoesNotContain("Public Function AddSix(", destinationRefactored);
        }

        [TestCase(MoveEndpoints.StdToStd, "Public", ThisStrategy)]
        [TestCase(MoveEndpoints.StdToStd, "Private", ThisStrategy)]
        [TestCase(MoveEndpoints.ClassToStd, "Public", null)]
        [TestCase(MoveEndpoints.FormToStd, "Public", null)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ExclusiveCallChainMemberExternallyReferences(MoveEndpoints endpoints, string accessibility, string expectedStrategy)
        {
            var classInstanceCode = string.Empty;
            var memberToMove = "Foo";
            var source =
$@"
Option Explicit

Private mfoo5 As Long, mfoo As Long, mfoo2 As Long, mfoo3 As Long, mfoo4 As Long

{accessibility} Function Foo(arg1 As Long) As Long
    mfoo = Bar(arg1)
    Foo = mfoo
End Function

Public Sub Goo(arg1 As Long)
    mfoo5 = AddSeven(arg1)
End Sub

Public Function AddSix(arg1 As Long) As Long
    AddSix = arg1 + 6
End Function

Private Function AddSeven(arg1 As Long) As Long
    AddSeven = arg1 + 7
End Function

Private Function Bar(arg1 As Long) As Long
    mfoo2 = arg1
    Bar = Barn(mfoo2)
End Function

Private Function Barn(arg1 As Long) As Long
    mfoo3 = arg1
    Barn = Bark(mfoo3)
End Function

Private Function Bark(arg1 As Long) As Long
    mfoo4 = AddSix(arg1)
    Bark = mfoo4
End Function
";

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType));

            var callSiteModuleName = "Module3";

            var memberAccessQualifier = moveDefinition.SourceModuleName;

            if (moveDefinition.IsClassSource || moveDefinition.IsFormSource)
            {
                memberAccessQualifier = "classInstance";
                classInstanceCode = $"{Support.ClassInstantiationBoilerPlate(memberAccessQualifier, callSiteModuleName)}";
            }

            var callSiteCode =
$@"
Option Explicit

Private mBar As Long

{classInstanceCode}

Public Sub MemberAccess(arg1 As Long)
    mBar = {memberAccessQualifier}.AddSix(arg1)
End Sub

Public Sub WithMemberAccess(arg2 As Long)
    With {memberAccessQualifier}
        mBar = .AddSix(arg2)
    End With
End Sub

Public Sub NonQualifiedAccess(arg3 As Long)
    mBar = AddSix(arg3)
End Sub
";
            moveDefinition.Add(new ModuleDefinition(callSiteModuleName, ComponentType.StandardModule, callSiteCode));

            var refactoredCode = RefactoredCode_UserSetsDestinationModuleName(moveDefinition, source);

            StringAssert.AreEqualIgnoringCase(ThisStrategy, refactoredCode.StrategyName);

            if (expectedStrategy is null)
            {
                return;
            }

            var sourceRefactored = refactoredCode.Source;
            StringAssert.DoesNotContain("Private Function Bar", sourceRefactored);
            StringAssert.DoesNotContain("Private Function Barn", sourceRefactored);
            StringAssert.DoesNotContain("Private Function Bark", sourceRefactored);
            StringAssert.DoesNotContain("Private mfoo As Long, mfoo2 As Long, mfoo3 As Long, mfoo4 As Long", sourceRefactored);
            StringAssert.Contains("Public Function AddSix(", sourceRefactored);
            StringAssert.Contains("Private Function AddSeven(", sourceRefactored);
            StringAssert.Contains("Private mfoo5 As Long", sourceRefactored);

            var destinationRefactored = refactoredCode.Destination;
            StringAssert.Contains("Private Function Bar", destinationRefactored);
            StringAssert.Contains("Private Function Barn", destinationRefactored);
            StringAssert.Contains("Private Function Bark", destinationRefactored);
            StringAssert.Contains($"{moveDefinition.SourceModuleName}.AddSix", destinationRefactored);
            StringAssert.Contains(" mfoo As Long", destinationRefactored);
            StringAssert.Contains(" mfoo2 As Long", destinationRefactored);

            var callSiteRefactored = refactoredCode[callSiteModuleName];
            StringAssert.AreEqualIgnoringCase(callSiteCode, callSiteRefactored);
        }

        [TestCase(MoveEndpoints.StdToStd)]
        [TestCase(MoveEndpoints.ClassToStd)]
        [TestCase(MoveEndpoints.FormToStd)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void RemovesMovedMemberScopeResolutionInDestination(MoveEndpoints endpoints)
        {
            var moveDefinition = new TestMoveDefinition(endpoints, ("Foo", ThisDeclarationType));

            var destinationModuleName = moveDefinition.DestinationModuleName;
            var source =
$@"
Option Explicit

Private mfoo As Long
Private mgoo As Long

Public Function Foo(arg1 As Long) As Long
    mfoo = arg1 * 10
    If {destinationModuleName}.LogIsEnabled Then
        {destinationModuleName}.Log ""Foo called""
        {destinationModuleName}.Entries = {destinationModuleName}.Entries + 1
    Endif
    Foo = mfoo
End Function

Public Property Let Goo(arg1 As Long)
    mgoo = arg1
    {destinationModuleName}.Log ""Let Goo called""
End Property

Public Property Get Goo() As Long
    Goo = mgoo
    {destinationModuleName}.Log ""Get Goo called""
End Property";


            var destination =
$@"
Option Explicit

Private Const LOG_IS_ENABLED = True

Public Entries As Long

Public Property Get LogIsEnabled()
    LogIsEnabled = LOG_IS_ENABLED
End Property

Public Sub Log(msg As String)
End Sub
";

            var refactorResults = RefactoredCode(moveDefinition, source, destination);

            var destinationExpectedContent =
                @"
Option Explicit

Private Const LOG_IS_ENABLED = True

Public Entries As Long

Public Function Foo(ByRef arg1 As Long)
    mfoo = arg1 * 10
    If LogIsEnabled Then
        Log ""Foo called""
        Entries = Entries + 1
    Endif
    Foo = mfoo
End Function

Public Property Get LogIsEnabled()
    LogIsEnabled = LOG_IS_ENABLED
End Property

Public Sub Log(msg As String)
End Sub
";
            var expectedLines = destinationExpectedContent.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var line in expectedLines)
            {
                StringAssert.Contains(line, refactorResults.Destination);
            }
        }

        [TestCase("Public", MoveEndpoints.StdToStd, ThisStrategy)]
        [TestCase("Private", MoveEndpoints.StdToStd, ThisStrategy)]
        [TestCase("Public", MoveEndpoints.ClassToStd, null)]
        [TestCase("Private", MoveEndpoints.ClassToStd, null)]
        [TestCase("Public", MoveEndpoints.FormToStd, null)]
        [TestCase("Private", MoveEndpoints.FormToStd, null)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ReferencesNonExclusiveProperty(string accessibility, MoveEndpoints endpoints, string expectedStrategy)
        {
            var memberToMove = "Foo";
            var source = $@"
Option Explicit

Private mBar As Long

{accessibility} Function Foo(arg1 As Long) As Long
    arg1 = Bar * 10
    Foo = arg1
End Function

Private Sub FooBar(arg1 As Long)
    Bar = arg1
End Sub

Public Property Let Bar(arg1 As Long)
    mBar = arg1
End Property

Public Property Get Bar() As Long
    Bar = mBar
End Property
";

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType));
            var refactoredCode = RefactoredCode_UserSetsDestinationModuleName(moveDefinition, source);

            StringAssert.AreEqualIgnoringCase(expectedStrategy, refactoredCode.StrategyName);
            if (expectedStrategy is null)
            {
                return;
            }

            StringAssert.DoesNotContain("Function Foo(arg1", refactoredCode.Source);
            StringAssert.Contains("Public Function Foo(ByRef arg1", refactoredCode.Destination);
            StringAssert.Contains($"arg1 = {moveDefinition.SourceModuleName}.Bar * 10", refactoredCode.Destination);
        }

        [TestCase(MoveEndpoints.StdToStd, ThisStrategy)]
        [TestCase(MoveEndpoints.ClassToStd, null)]
        [TestCase(MoveEndpoints.FormToStd, null)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void SupportMemberReferencesNonExclusiveBackingVariable( MoveEndpoints endpoints, string expectedStrategy)
        {
            var memberToMove = "Foo";
            var source = $@"
Option Explicit

Private mBar As Long

Public Sub InitializeModule(arg1 As Long)
    mBar = arg1
End Sub

Public Function Foo(arg1 As Long) As Long
    arg1 = Bar * 10
    Bar = arg1
    Foo = arg1
End Function

Public Property Let Bar(arg1 As Long)
    mBar = arg1
End Property

Public Property Get Bar() As Long
    Bar = mBar
End Property
";

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType));
            var refactoredCode = RefactoredCode_UserSetsDestinationModuleName(moveDefinition, source);

            StringAssert.AreEqualIgnoringCase(expectedStrategy, refactoredCode.StrategyName);
            if (expectedStrategy is null)
            {
                return;
            }

            StringAssert.Contains("Bar(", refactoredCode.Source);
            StringAssert.DoesNotContain("Function Foo(arg1", refactoredCode.Source);
            StringAssert.Contains("Public Function Foo(ByRef arg1", refactoredCode.Destination);
            StringAssert.Contains($"arg1 = {moveDefinition.SourceModuleName}.Bar", refactoredCode.Destination);
            StringAssert.Contains($"{moveDefinition.SourceModuleName}.Bar = arg1", refactoredCode.Destination);
            StringAssert.DoesNotContain("Bar(", refactoredCode.Destination);
        }

        [TestCase("Public", MoveEndpoints.StdToStd)]
        [TestCase("Private", MoveEndpoints.StdToStd)]
        [TestCase("Public", MoveEndpoints.ClassToStd)]
        [TestCase("Private", MoveEndpoints.ClassToStd)]
        [TestCase("Public", MoveEndpoints.FormToStd)]
        [TestCase("Private", MoveEndpoints.FormToStd)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ReferencesExclusivePropertyAndBackingVariable(string accessibility, MoveEndpoints endpoints)
        {
            var memberToMove = "Foo";
            var source = $@"
Option Explicit

Private mBar As Long

{accessibility} Function Foo(arg1 As Long) As Long
    arg1 = Bar * 10
    Bar = arg1
    Foo = arg1
End Function

Public Property Let Bar(arg1 As Long)
    mBar = arg1
End Property

Public Property Get Bar() As Long
    Bar = mBar
End Property
";

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType));
            var refactoredCode = RefactoredCode_UserSetsDestinationModuleName(moveDefinition, source);

            StringAssert.AreEqualIgnoringCase(ThisStrategy, refactoredCode.StrategyName);

            StringAssert.DoesNotContain("Function Foo(arg1", refactoredCode.Source);
            StringAssert.DoesNotContain("Public Property Let Bar", refactoredCode.Source);
            StringAssert.DoesNotContain("Public Property Get Bar", refactoredCode.Source);

            StringAssert.Contains("Public Function Foo(ByRef arg1", refactoredCode.Destination);
            StringAssert.Contains($"arg1 = Bar * 10", refactoredCode.Destination);
            StringAssert.Contains($"Bar = arg1", refactoredCode.Destination);
        }

        [TestCase("Public", MoveEndpoints.StdToStd)]
        [TestCase("Private", MoveEndpoints.StdToStd)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ReferencesExternallyReferencedSupportProperty(string accessibility, MoveEndpoints endpoints)
        {
            var memberToMove = "Foo";
            var source = $@"
Option Explicit

Private mBar As Long

{accessibility} Function Foo(arg1 As Long) As Long
    arg1 = Bar * 10
    Bar = arg1
    Foo = arg1
End Function

Public Property Let Bar(arg1 As Long)
    mBar = arg1
End Property

Public Property Get Bar() As Long
    Bar = mBar
End Property
";

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType));

            var externalReferencingCode =
$@"
Public Sub FooBar(arg1 As Long)
    {moveDefinition.SourceModuleName}.Bar = arg1
End Sub
";
            moveDefinition.Add(new ModuleDefinition("Module3", ComponentType.StandardModule, externalReferencingCode));

            var refactoredCode = RefactoredCode_UserSetsDestinationModuleName(moveDefinition, source);

            StringAssert.AreEqualIgnoringCase(ThisStrategy, refactoredCode.StrategyName);

            StringAssert.DoesNotContain("Function Foo(arg1", refactoredCode.Source);

            StringAssert.Contains("Public Function Foo(ByRef arg1", refactoredCode.Destination);
            StringAssert.DoesNotContain("Public Property Let Bar", refactoredCode.Destination);
            StringAssert.DoesNotContain("Public Property Get Bar", refactoredCode.Destination);

            var module3Content = refactoredCode["Module3"];
            StringAssert.Contains($"{moveDefinition.SourceModuleName}.Bar", refactoredCode.Destination);

        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void CorrectsMemberNameCollisionInDestination()
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

Public multiplier As Long

Public Function Goo(arg1 As Long) 
    Goo = multiplier * arg1
End Function
";

            var refactorResults = RefactoredCode(moveDefinition, source, destination);

            var destinationExpectedContent =
$@"
Option Explicit

Public multiplier As Long

Private mgoo As Long

Public Property Let Goo1(ByVal arg1 As Long)
    mgoo = arg1
End Property

Public Property Get Goo1() As Long
    Goo1 = mgoo
End Property

Public Function Goo(arg1 As Long) 
    Goo = multiplier * arg1
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
        public void SetsNewMemberNameAtExternalReferences()
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

Public multiplier As Long

Public Function Goo(arg1 As Long) 
    Goo = multiplier * arg1
End Function
";
            var callSiteModuleName = "Module3";
            var callSiteCode =
    $@"
Option Explicit

Private mBar As Long

Public Sub MemberAccess(arg1 As Long)
    mBar = {moveDefinition.SourceModuleName}.Goo
End Sub

Public Sub WithMemberAccess(arg2 As Long)
    With {moveDefinition.SourceModuleName}
        mBar = .Goo
    End With
End Sub

Public Sub NonQualified(arg3 As Long)
    mBar = Goo + arg3
End Sub
";
            moveDefinition.Add(new ModuleDefinition(callSiteModuleName, ComponentType.StandardModule, callSiteCode));

            var refactorResults = RefactoredCode(moveDefinition, source, destination);

            var callSiteExpectedContent =
$@"
Option Explicit

Private mBar As Long

Public Sub MemberAccess(arg1 As Long)
    mBar = {moveDefinition.DestinationModuleName}.Goo1
End Sub

Public Sub WithMemberAccess(arg2 As Long)
    With {moveDefinition.SourceModuleName}
        mBar = {moveDefinition.DestinationModuleName}.Goo1
    End With
End Sub

Public Sub NonQualified(arg3 As Long)
    mBar = {moveDefinition.DestinationModuleName}.Goo1 + arg3
End Sub
";
            var expectedLines = callSiteExpectedContent.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var line in expectedLines)
            {
                StringAssert.Contains(line, refactorResults[callSiteModuleName]);
            }
        }


        [TestCase(MoveEndpoints.StdToStd)]
        [TestCase(MoveEndpoints.ClassToStd)]
        [TestCase(MoveEndpoints.FormToStd)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ContainsObjectField(MoveEndpoints endpoints)
        {
            var memberToMove = "FooMath";
            var source =
$@"
Option Explicit

Private mObj As ObjectClass

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

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType));

            var refactoredCode = RefactoredCode_UserSetsDestinationModuleName(moveDefinition, source);

            StringAssert.AreEqualIgnoringCase(ThisStrategy, refactoredCode.StrategyName);

            var sourceRefactored = refactoredCode.Source;
            StringAssert.DoesNotContain("FooMath", sourceRefactored);
            StringAssert.DoesNotContain("mObj", sourceRefactored);

            var destinationRefactored = refactoredCode.Destination;
            StringAssert.Contains("FooMath", destinationRefactored);
            StringAssert.Contains("mObj", destinationRefactored);
        }
    }
}
