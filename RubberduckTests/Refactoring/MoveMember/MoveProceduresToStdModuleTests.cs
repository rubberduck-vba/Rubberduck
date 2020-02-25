using NUnit.Framework;
using Rubberduck.Common;
using Rubberduck.Parsing.Grammar;
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
    [TestFixture]
    public class MoveProceduresToStdModuleTests : MoveMemberTestsBase
    {
        private const string ThisStrategy = nameof(MoveMemberToStdModule);
        private const DeclarationType ThisDeclarationType = DeclarationType.Procedure;

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
            var source = $@"
Option Explicit

{accessibility} Sub Log()
End Sub";

            var moveDefinition = new TestMoveDefinition(endpoints, ("Log", ThisDeclarationType));

            var refactoredCode = RefactoredCode_UserSetsDestinationModuleName(moveDefinition, source);

            Assert.AreEqual(null, refactoredCode.StrategyName);
        }

        [TestCase(MoveEndpoints.StdToStd)]
        [TestCase(MoveEndpoints.ClassToStd)]
        [TestCase(MoveEndpoints.FormToStd)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void PreviewMovedContent(MoveEndpoints endpoints)
        {
            var memberToMove = "Foo";
            var source =
$@"
Option Explicit

Sub Foo(ByVal arg1 As Long, ByRef result As Long)
    result = 10 * arg1
End Sub
";

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType));

            var preview = RetrievePreviewAfterUserInput(moveDefinition, source, (memberToMove, ThisDeclarationType));

            StringAssert.Contains("Option Explicit", preview);
            Assert.IsTrue(Support.OccursOnce("Public Sub Foo(", preview));
        }

        [TestCase(MoveEndpoints.StdToStd, "Public", ThisStrategy)]
        [TestCase(MoveEndpoints.StdToStd, "Private", ThisStrategy)]
        [TestCase(MoveEndpoints.ClassToStd, "Public", ThisStrategy)]
        [TestCase(MoveEndpoints.ClassToStd, "Private", ThisStrategy)]
        [TestCase(MoveEndpoints.FormToStd, "Public", ThisStrategy)]
        [TestCase(MoveEndpoints.FormToStd, "Private", ThisStrategy)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MovedSubReferencesExclusiveSupportConstant(MoveEndpoints endpoints, string exclusiveFuncAccessibility, string expectedStrategy)
        {
            var memberToMove = "CalculateVolumeFromDiameter";
            var exclusiveSupportElement = "Pi";
            var source =
$@"
Option Explicit

Private Const {exclusiveSupportElement} As Single = 3.14

Public Sub CalculateVolumeFromDiameter(ByVal diameter As Single, ByVal height As Single, ByRef volume As Single)
    CalculateVolume(diameter / 2, height, volume)
End Sub

{exclusiveFuncAccessibility} Sub CalculateVolume(ByVal radius As Single, ByVal height As Single, ByRef volume As Single)
    volume = height * {exclusiveSupportElement} * radius ^ 2
End Sub
";

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType));

            var refactoredCode = RefactoredCode_UserSetsDestinationModuleName(moveDefinition, source);

            StringAssert.AreEqualIgnoringCase(expectedStrategy, refactoredCode.StrategyName);

            if (expectedStrategy is null)
            {
                return;
            }

            StringAssert.DoesNotContain("Public Sub CalculateVolumeFromDiameter(", refactoredCode.Source);
            StringAssert.DoesNotContain($"{exclusiveFuncAccessibility} Function CalculateVolume(", refactoredCode.Source);
            StringAssert.DoesNotContain($"{exclusiveSupportElement} As Single", refactoredCode.Source);

            StringAssert.Contains($"Private Const {exclusiveSupportElement} As Single", refactoredCode.Destination);
            StringAssert.Contains("Public Sub CalculateVolumeFromDiameter(", refactoredCode.Destination);
            StringAssert.Contains("CalculateVolume(diameter / 2, height, volume)", refactoredCode.Destination);
            StringAssert.Contains($"{exclusiveFuncAccessibility} Sub CalculateVolume(", refactoredCode.Destination);
        }

        [TestCase(MoveEndpoints.StdToStd, ThisStrategy)]
        [TestCase(MoveEndpoints.ClassToStd, null)]
        [TestCase(MoveEndpoints.FormToStd, null)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MovedSubReferencesNonExclusivePublicSupportConstant(MoveEndpoints endpoints, string expectedStrategy)
        {
            var memberToMove = "CalculateVolumeFromDiameter";
            var source =
$@"
Option Explicit

Public Const Pi As Single = 3.14

Public Sub CalculateVolumeFromDiameter(ByVal diameter As Single, ByVal height As Single, ByRef volume As Single)
    volume = height * Pi * (diameter / 2) ^ 2 
End Sub

Public Function CalculateCircumferenceFromDiameter(ByVal diameter As Single) As Single
    CalculateCircumferenceFromDiameter = diameter * Pi
End Function";

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType));

            var refactoredCode = RefactoredCode_UserSetsDestinationModuleName(moveDefinition, source);

            StringAssert.AreEqualIgnoringCase(expectedStrategy, refactoredCode.StrategyName);

            if (expectedStrategy is null)
            {
                return;
            }

            StringAssert.DoesNotContain("Public Sub CalculateVolumeFromDiameter(", refactoredCode.Source);
            StringAssert.Contains($"Public Function CalculateCircumferenceFromDiameter(", refactoredCode.Source);

            StringAssert.Contains($"Sub CalculateVolumeFromDiameter(", refactoredCode.Destination);
            StringAssert.Contains($"height * {moveDefinition.SourceModuleName}.Pi", refactoredCode.Destination);
        }

        [TestCase(MoveEndpoints.StdToStd, null)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MovedSubReferencesNonExclusivePrivateSupportConstant(MoveEndpoints endpoints, string expectedStrategy)
        {
            var memberToMove = "CalculateVolumeFromDiameter";
            var pi = "Pi";
            var source =
$@"
Option Explicit

Private Const {pi} As Single = 3.14

Public Sub CalculateVolumeFromDiameter(ByVal diameter As Single, ByVal height As Single, ByRef volume As Single)
    volume = height * {pi} * (diameter / 2) ^ 2 
End Sub

Public Sub CalculateCircumferenceFromDiameter(ByVal diameter As Single, ByRef circumference As Single)
    circumference = diameter * {pi}
End Sub";

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

{accessibility} Sub Foo(arg1 As Long)
    bar = bar + arg1
End Sub

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

            StringAssert.DoesNotContain($"{accessibility} Sub Foo(", refactoredCode.Source);
            StringAssert.Contains($"Public Sub Goo(", refactoredCode.Source);
            StringAssert.Contains($"Public bar As Long", refactoredCode.Source);

            StringAssert.Contains($"Public Sub Foo(", refactoredCode.Destination);
            StringAssert.Contains($"{moveDefinition.SourceModuleName}.bar = {moveDefinition.SourceModuleName}.bar + arg1", refactoredCode.Destination);
            StringAssert.DoesNotContain($"Public bar As Long", refactoredCode.Destination);
        }

        [TestCase(MoveEndpoints.StdToStd, null)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ReferencesNonExclusivePrivateField(MoveEndpoints endpoints, string expectedStrategy)
        {
            var memberToMove = "Foo";
            var source = $@"
Option Explicit

Private bar As Long

Public Sub Foo(arg1 As Long)
    bar = bar + arg1
End Sub

Public Sub Goo(arg1 As Long)
    bar = bar + arg1
End Sub
";

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType));

            var refactoredCode = RefactoredCode_UserSetsDestinationModuleName(moveDefinition, source);

            StringAssert.AreEqualIgnoringCase(expectedStrategy, refactoredCode.StrategyName);
        }

        [TestCase("Public", MoveEndpoints.StdToStd, ThisStrategy)]
        [TestCase("Private", MoveEndpoints.StdToStd, null)]
        [TestCase("Public", MoveEndpoints.ClassToStd, null)]
        [TestCase("Private", MoveEndpoints.ClassToStd, null)]
        [TestCase("Public", MoveEndpoints.FormToStd, null)]
        [TestCase("Private", MoveEndpoints.FormToStd, null)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ReferencesNonExclusiveMember(string subLogAccessibility, MoveEndpoints endpoints, string expectedStrategyName)
        {
            var memberToMove = "Foo";
            var source = $@"
Option Explicit

Public Sub Foo(arg1 As Long)
    Log
End Sub

Public Sub Goo(arg1 As Long)
    Log
End Sub

{subLogAccessibility} Sub Log()
End Sub";

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType));
            var refactoredCode = RefactoredCode_UserSetsDestinationModuleName(moveDefinition, source);

            if (expectedStrategyName is null)
            {
                return;
            }

            StringAssert.DoesNotContain("Public Sub Foo(arg1 As Long)", refactoredCode.Source);

            StringAssert.Contains("Public Sub Foo(ByRef arg1 As Long)", refactoredCode.Destination);
            StringAssert.Contains($"{moveDefinition.SourceModuleName}.Log", refactoredCode.Destination);
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

Public Sub CalculateVolumeFromDiameter(ByVal diameter As Single, ByVal height As Single, ByRef volume As Single)
    volume = height * 3.14 * (diameter / 2) ^ 2 
End Sub
";


            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, (memberToMove, ThisDeclarationType));

            var externalReferences =
$@"
Option Explicit

Private mFoo As Single

Public Sub MemberAccess()
    {moveDefinition.SourceModuleName}.CalculateVolumeFromDiameter 7.5, 4.2, mFoo
End Sub

Public Sub WithMemberAccess()
    With {moveDefinition.SourceModuleName}
        .CalculateVolumeFromDiameter 8.5, 4.2, mFoo
    End With
End Sub

Public Sub NonQualifiedAccess()
    CalculateVolumeFromDiameter 9.5, 4.2, mFoo
End Sub
";
            var referencingModuleName = "Module3";
            moveDefinition.Add(new ModuleDefinition(referencingModuleName, ComponentType.StandardModule, externalReferences));

            var refactoredCode = RefactoredCode_UserSetsDestinationModuleName(moveDefinition, source);

            StringAssert.AreEqualIgnoringCase(ThisStrategy, refactoredCode.StrategyName);

            StringAssert.DoesNotContain("Public Sub CalculateVolumeFromDiameter(", refactoredCode.Source);
            StringAssert.Contains($"Public Sub CalculateVolumeFromDiameter(", refactoredCode.Destination);

            var module3Content = refactoredCode[referencingModuleName];

            StringAssert.Contains($"{moveDefinition.DestinationModuleName}.CalculateVolumeFromDiameter 7.5, 4.2, mFoo", module3Content);
            StringAssert.Contains($"With {moveDefinition.SourceModuleName}", module3Content);
            StringAssert.Contains($"{moveDefinition.DestinationModuleName}.CalculateVolumeFromDiameter 8.5, 4.2, mFoo", module3Content);
            StringAssert.Contains($"{moveDefinition.DestinationModuleName}.CalculateVolumeFromDiameter 9.5, 4.2, mFoo", module3Content);
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

Public Sub CalculateVolumeFromDiameter(ByVal diameter As Single, ByVal height As Single, ByRef volume As Single)
    volume = height * 3.14 * (diameter / 2) ^ 2 
End Sub
";

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType));

            var instanceIdentifier = "mClassOrUserForm";
            var externalReferences =
$@"
Option Explicit

Private mFoo As Single

{Support.ClassInstantiationBoilerPlate(instanceIdentifier, moveDefinition.SourceModuleName)}

Public Sub MemberAccess()
    {instanceIdentifier}.CalculateVolumeFromDiameter 7.5, 4.2, mFoo
End Sub

Public Sub WithMemberAccess()
    With {instanceIdentifier}
        .CalculateVolumeFromDiameter 8.5, 4.2, mFoo
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
            var memberToMove = "RenameFile";
            var source =
$@"
Option Explicit

Private mfileName As String

Public Sub ChangeLogFile(fileName As String)
    RenameFile mfileName, fileName
    mfileName = fileName
End Sub

{targetAccessibility} Sub RenameFile(oldName As String, newName As String)
    Name oldName As newName
End Sub
";

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType));

            var refactoredCode = RefactoredCode_UserSetsDestinationModuleName(moveDefinition, source);

            StringAssert.AreEqualIgnoringCase(ThisStrategy, refactoredCode.StrategyName);

            var sourceRefactored = refactoredCode.Source;
            StringAssert.Contains($"{moveDefinition.DestinationModuleName}.RenameFile mfileName, fileName", sourceRefactored);
            StringAssert.DoesNotContain($"{targetAccessibility} Sub RenameFile", sourceRefactored);

            StringAssert.Contains("Public Sub RenameFile", refactoredCode.Destination);
        }


        [TestCase(MoveEndpoints.StdToStd)]
        [TestCase(MoveEndpoints.ClassToStd)]
        [TestCase(MoveEndpoints.FormToStd)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void InternallyReferencedSubToStdModule(MoveEndpoints endpoints)
        {
            var memberToMove = "CalculateVolume";

            var source =
$@"
Option Explicit

Private Const Pi As Single = 3.14

Public Sub CalculateCylinderVolumeFromDiameter(diameter As Single, height As Single, ByRef volume As Single)
    CalculateVolume(Pi * (diameter / 2) ^ 2, height, volume)
End Sub

Public Sub CalculateCylinderVolumeFromRadius(radius As Single, height As Single, ByRef circumference As Single)
    CalculateVolume(Pi * (radius) ^ 2, height, volume)
End Sub

Private Sub CalculateVolume(area As Single, height As Single, volume As Single)
    volume = area * height
End Sub
";

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType));

            var refactoredCode = RefactoredCode_UserSetsDestinationModuleName(moveDefinition, source);

            StringAssert.AreEqualIgnoringCase(ThisStrategy, refactoredCode.StrategyName);
            StringAssert.Contains($"{moveDefinition.DestinationModuleName}.CalculateVolume(", refactoredCode.Source);
            StringAssert.DoesNotContain("Private Function CalculateVolume", refactoredCode.Source);

            StringAssert.Contains($"volume = area * height", refactoredCode.Destination);
            StringAssert.Contains($"Public Sub CalculateVolume(ByRef area", refactoredCode.Destination);
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

{accessibility} Sub Foo(arg1 As Long)
    mfoo = Bar(arg1)
End Sub

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

{accessibility} Sub Foo(arg1 As Long)
    mfoo = Bar(arg1)
End Sub

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
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void UserSelectsAdditionalMethodToEnableMove(MoveEndpoints endpoints, string accessibility)
        {
            var memberToMove = "Foo";
            var mfoo5 = "mfoo5";
            var source =
$@"
Option Explicit

Private {mfoo5} As Long, mfoo As Long, mfoo2 As Long, mfoo3 As Long, mfoo4 As Long

{accessibility} Sub Foo(arg1 As Long)
    mfoo = Bar(arg1)
End Sub

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

            var refactoredCode = RefactoredCode(moveDefinition, source, null, null, false, ("Goo", DeclarationType.Procedure));

            StringAssert.AreEqualIgnoringCase(ThisStrategy, refactoredCode.StrategyName);

            var sourceRefactored = refactoredCode.Source.Trim();
            StringAssert.AreEqualIgnoringCase("Option Explicit", sourceRefactored);

            var destinationRefactored = refactoredCode.Destination;
            StringAssert.Contains("Private Function Bar", destinationRefactored);
            StringAssert.Contains("Private Function Barn", destinationRefactored);
            StringAssert.Contains("Private Function Bark", destinationRefactored);
            StringAssert.Contains(" mfoo As Long", destinationRefactored);
            StringAssert.Contains(" mfoo2 As Long", destinationRefactored);
            StringAssert.Contains(" mfoo5 As Long", destinationRefactored);
            StringAssert.Contains("Private Function AddSix(", destinationRefactored);
            StringAssert.Contains("Public Sub Goo(", destinationRefactored);
            StringAssert.Contains("Private Function AddSeven(", destinationRefactored);
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

{accessibility} Sub Foo(arg1 As Long)
    mfoo = Bar(arg1)
End Sub

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

Public Sub Foo(arg1 As Long)
    mfoo = Bar(arg1)
End Sub

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

{accessibility} Sub Foo(arg1 As Long)
    mfoo = Bar(arg1)
End Sub

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
        public void RemovesMovedSubScopeResolutionInDestination(MoveEndpoints endpoints)
        {
            var moveDefinition = new TestMoveDefinition(endpoints, ("Foo", ThisDeclarationType));

            var destinationModuleName = moveDefinition.DestinationModuleName;
            var source =
$@"
Option Explicit

Private mfoo As Long
Private mgoo As Long

Public Sub Foo(arg1 As Long)
    mfoo = arg1
    If {destinationModuleName}.LogIsEnabled Then
        {destinationModuleName}.Log ""Foo called""
        {destinationModuleName}.Entries = {destinationModuleName}.Entries + 1
    Endif
End Sub

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

Public Sub Foo(ByRef arg1 As Long)
    mfoo = arg1
    If LogIsEnabled Then
        Log ""Foo called""
        Entries = Entries + 1
    Endif
End Sub

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

{accessibility} Sub Foo(arg1 As Long)
    arg1 = Bar
End Sub

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

            StringAssert.DoesNotContain("Sub Foo(arg1", refactoredCode.Source);
            StringAssert.Contains("Public Sub Foo(ByRef arg1", refactoredCode.Destination);
            StringAssert.Contains($"arg1 = {moveDefinition.SourceModuleName}", refactoredCode.Destination);
        }

        [TestCase(MoveEndpoints.StdToStd, ThisStrategy)]
        [TestCase(MoveEndpoints.ClassToStd, null)]
        [TestCase(MoveEndpoints.FormToStd, null)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void SupportMembersReferenceNonExclusiveBackingVariables(MoveEndpoints endpoints, string expectedStrategy)
        {
            var memberToMove = "Foo";
            var source = $@"
Option Explicit

Private mBar As Long

Public Sub InitializeModule(arg1 As Long)
    mBar = arg1
End Sub

Public Sub Foo(arg1 As Long)
    arg1 = Bar * 10
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

            StringAssert.Contains("Bar(", refactoredCode.Source);
            StringAssert.DoesNotContain("Sub Foo(arg1", refactoredCode.Source);
            StringAssert.Contains("Public Sub Foo(ByRef arg1", refactoredCode.Destination);
            StringAssert.Contains($"arg1 = {moveDefinition.SourceModuleName}.Bar", refactoredCode.Destination);
            StringAssert.Contains($"{moveDefinition.SourceModuleName}.Bar = arg1", refactoredCode.Destination);
            StringAssert.DoesNotContain("Bar(", refactoredCode.Destination);
        }


        [TestCase(MoveEndpoints.StdToStd, ThisStrategy)]
        [TestCase(MoveEndpoints.ClassToStd, null)]
        [TestCase(MoveEndpoints.FormToStd, null)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void SupportMembersReferenceNonExclusivePrivateMember(MoveEndpoints endpoints, string expectedStrategy)
        {
            var memberToMove = "Foo";
            var source = $@"
Option Explicit

Private mBar As Long

Public Sub InitializeModule(arg1 As Long)
    PointlessSub arg1
End Sub

Private Sub PointlessSub(arg1 As Long)
    mBar = arg1
End Sub

Public Sub Foo(arg1 As Long)
    arg1 = Bar * 10
    Bar = arg1
End Sub

Public Property Let Bar(arg1 As Long)
    PointlessSub arg1
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
            StringAssert.DoesNotContain("Sub Foo(arg1", refactoredCode.Source);
            StringAssert.Contains("Public Sub Foo(ByRef arg1", refactoredCode.Destination);
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

{accessibility} Sub Foo(arg1 As Long)
    arg1 = Bar * 10
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

            StringAssert.AreEqualIgnoringCase(ThisStrategy, refactoredCode.StrategyName);

            StringAssert.DoesNotContain("Sub Foo(arg1", refactoredCode.Source);
            StringAssert.DoesNotContain("Public Property Let Bar", refactoredCode.Source);
            StringAssert.DoesNotContain("Public Property Get Bar", refactoredCode.Source);

            StringAssert.Contains("Public Sub Foo(ByRef arg1", refactoredCode.Destination);
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

{accessibility} Sub Foo(arg1 As Long)
    arg1 = Bar * 10
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

            var externalReferencingCode =
$@"
Public Sub FooBar(arg1 As Long)
    {moveDefinition.SourceModuleName}.Bar = arg1
End Sub
";
            moveDefinition.Add(new ModuleDefinition("Module3", ComponentType.StandardModule, externalReferencingCode));

            var refactoredCode = RefactoredCode_UserSetsDestinationModuleName(moveDefinition, source);

            StringAssert.AreEqualIgnoringCase(ThisStrategy, refactoredCode.StrategyName);

            StringAssert.DoesNotContain("Sub Foo(arg1", refactoredCode.Source);

            StringAssert.Contains("Public Sub Foo(ByRef arg1", refactoredCode.Destination);
            StringAssert.DoesNotContain("Public Property Let Bar", refactoredCode.Destination);
            StringAssert.DoesNotContain("Public Property Get Bar", refactoredCode.Destination);

            var module3Content = refactoredCode["Module3"];
            StringAssert.Contains($"{moveDefinition.SourceModuleName}.Bar", refactoredCode.Destination);

        }
    }
}
