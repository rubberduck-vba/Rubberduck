using NUnit.Framework;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.VBEditor.SafeComWrappers;
using System;


namespace RubberduckTests.Refactoring.MoveMember
{
    [TestFixture]
    public class MoveProceduresToStdModuleTests : MoveMemberRefactoringActionTestSupportBase
    {
        [TestCase(MoveEndpoints.StdToClass)]
        [TestCase(MoveEndpoints.ClassToClass)]
        [TestCase(MoveEndpoints.FormToClass)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void SimpleMoveToClassModuleMoveThrow(MoveEndpoints endpoints)
        {
            var source = $@"
Option Explicit

Public Sub Log()
End Sub";

            ExecuteSingleTargetMoveThrowsExceptionTest(("Log", DeclarationType.Procedure), endpoints, source);
        }

        [TestCase(MoveEndpoints.StdToStd, "Public")]
        [TestCase(MoveEndpoints.StdToStd, "Private")]
        [TestCase(MoveEndpoints.ClassToStd, "Public")]
        [TestCase(MoveEndpoints.ClassToStd, "Private")]
        [TestCase(MoveEndpoints.FormToStd, "Public")]
        [TestCase(MoveEndpoints.FormToStd, "Private")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void SupportSubReferencesExclusiveSupportConstant(MoveEndpoints endpoints, string exclusiveFuncAccessibility)
        {
            var memberToMove = ("CalculateVolumeFromDiameter", DeclarationType.Procedure);
            var exclusiveSupportConstant = "Pi";
            var source =
$@"
Option Explicit

Private Const {exclusiveSupportConstant} As Single = 3.14

Public Sub CalculateVolumeFromDiameter(ByVal diameter As Single, ByVal height As Single, ByRef volume As Single)
    CalculateVolume(diameter / 2, height, volume)
End Sub

{exclusiveFuncAccessibility} Sub CalculateVolume(ByVal radius As Single, ByVal height As Single, ByRef volume As Single)
    volume = height * {exclusiveSupportConstant} * radius ^ 2
End Sub
";
            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source);

            if (endpoints.IsStdModuleSource())
            {
                if (exclusiveFuncAccessibility.Equals(Tokens.Private))
                {
                    StringAssert.DoesNotContain($"Private Const {exclusiveSupportConstant} As Single", refactoredCode.Source);
                    StringAssert.DoesNotContain($"{exclusiveFuncAccessibility} Sub CalculateVolume(", refactoredCode.Source);
                    StringAssert.DoesNotContain("Public Sub CalculateVolumeFromDiameter(", refactoredCode.Source);

                    StringAssert.Contains("Public Sub CalculateVolumeFromDiameter(", refactoredCode.Destination);
                    StringAssert.Contains("CalculateVolume(diameter / 2, height, volume)", refactoredCode.Destination);
                    StringAssert.Contains($"{exclusiveFuncAccessibility} Sub CalculateVolume(", refactoredCode.Destination);
                    StringAssert.Contains($"Private Const {exclusiveSupportConstant} As Single", refactoredCode.Destination);
                }
                else
                {
                    StringAssert.Contains($"Private Const {exclusiveSupportConstant} As Single", refactoredCode.Source);
                    StringAssert.Contains($"{exclusiveFuncAccessibility} Sub CalculateVolume(", refactoredCode.Source);
                    StringAssert.DoesNotContain("Public Sub CalculateVolumeFromDiameter(", refactoredCode.Source);

                    StringAssert.Contains("Public Sub CalculateVolumeFromDiameter(", refactoredCode.Destination);
                    StringAssert.Contains($"{endpoints.SourceModuleName()}.CalculateVolume(diameter / 2, height, volume)", refactoredCode.Destination);
                    StringAssert.DoesNotContain($"{exclusiveFuncAccessibility} Sub CalculateVolume(", refactoredCode.Destination);
                }
            }
            else
            {
                StringAssert.DoesNotContain($"Private Const {exclusiveSupportConstant} As Single", refactoredCode.Source);
                StringAssert.DoesNotContain($"{exclusiveFuncAccessibility} Sub CalculateVolume(", refactoredCode.Source);
                StringAssert.DoesNotContain("Public Sub CalculateVolumeFromDiameter(", refactoredCode.Source);

                StringAssert.Contains("Public Sub CalculateVolumeFromDiameter(", refactoredCode.Destination);
                StringAssert.Contains("CalculateVolume(diameter / 2, height, volume)", refactoredCode.Destination);
                StringAssert.Contains($"{exclusiveFuncAccessibility} Sub CalculateVolume(", refactoredCode.Destination);
                StringAssert.Contains($"Private Const {exclusiveSupportConstant} As Single", refactoredCode.Destination);
            }
        }

        [TestCase(MoveEndpoints.StdToStd, "Public Const Pi As Single = 3.14", false)]
        [TestCase(MoveEndpoints.StdToStd, "Public Pi As Single", false)]
        [TestCase(MoveEndpoints.ClassToStd, "Public Pi As Single", true)]
        [TestCase(MoveEndpoints.FormToStd, "Public Pi As Single", true)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MovedSubReferencesNonExclusivePublicSupportNonMember(MoveEndpoints endpoints, string nonMemberDeclaration, bool throwsException)
        {
            var memberToMove = ("CalculateVolumeFromDiameter", DeclarationType.Procedure);
            var source =
$@"
Option Explicit

{nonMemberDeclaration}

Public Sub CalculateVolumeFromDiameter(ByVal diameter As Single, ByVal height As Single, ByRef volume As Single)
    volume = height * Pi * (diameter / 2) ^ 2 
End Sub

Public Function CalculateCircumferenceFromDiameter(ByVal diameter As Single) As Single
    CalculateCircumferenceFromDiameter = diameter * Pi
End Function";
            if (throwsException)
            {
                ExecuteSingleTargetMoveThrowsExceptionTest(memberToMove, endpoints, source);
                return;
            }

            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source);

            StringAssert.DoesNotContain("Public Sub CalculateVolumeFromDiameter(", refactoredCode.Source);
            StringAssert.Contains($"Public Function CalculateCircumferenceFromDiameter(", refactoredCode.Source);

            StringAssert.Contains($"Sub CalculateVolumeFromDiameter(", refactoredCode.Destination);
            StringAssert.Contains($"height * {endpoints.SourceModuleName()}.Pi", refactoredCode.Destination);
        }

        [TestCase(MoveEndpoints.StdToStd, "Private Const Pi As Single = 3.14")]
        [TestCase(MoveEndpoints.StdToStd, "Private Pi As Single")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ReferencesNonExclusivePrivateSupportNonMember(MoveEndpoints endpoints, string nonMemberDeclaration)
        {
            var memberToMove = ("CalculateVolumeFromDiameter", DeclarationType.Procedure);
            var source =
$@"
Option Explicit

{nonMemberDeclaration}

Public Sub CalculateVolumeFromDiameter(ByVal diameter As Single, ByVal height As Single, ByRef volume As Single)
    volume = height * Pi * (diameter / 2) ^ 2 
End Sub

Public Sub CalculateCircumferenceFromDiameter(ByVal diameter As Single, ByRef circumference As Single)
    circumference = diameter * Pi
End Sub";

            ExecuteSingleTargetMoveThrowsExceptionTest(memberToMove, endpoints, source);
        }

        [TestCase("Public", MoveEndpoints.StdToStd, false)]
        [TestCase("Private", MoveEndpoints.StdToStd, true)]
        [TestCase("Public", MoveEndpoints.ClassToStd, true)]
        [TestCase("Private", MoveEndpoints.ClassToStd, true)]
        [TestCase("Public", MoveEndpoints.FormToStd, true)]
        [TestCase("Private", MoveEndpoints.FormToStd, true)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ReferencesNonExclusiveMember(string subLogAccessibility, MoveEndpoints endpoints, bool throwsException)
        {
            var memberToMove = ("Foo", DeclarationType.Procedure);
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

            if (throwsException)
            {
                ExecuteSingleTargetMoveThrowsExceptionTest(memberToMove, endpoints, source);
                return;
            }

            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source);

            StringAssert.DoesNotContain("Public Sub Foo(arg1 As Long)", refactoredCode.Source);
            StringAssert.Contains("Public Sub Foo(arg1 As Long)", refactoredCode.Destination);
            StringAssert.Contains($"{endpoints.SourceModuleName()}.Log", refactoredCode.Destination);
        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ExternallyReferencedMemberStdToStd()
        {
            var memberToMove = ("CalculateVolumeFromDiameter", DeclarationType.Procedure);
            var endpoints = MoveEndpoints.StdToStd;
            var source =
$@"
Option Explicit

Public Sub CalculateVolumeFromDiameter(ByVal diameter As Single, ByVal height As Single, ByRef volume As Single)
    volume = height * 3.14 * (diameter / 2) ^ 2 
End Sub
";

            var externalReferences =
$@"
Option Explicit

Private mFoo As Single

Public Sub MemberAccess()
    {endpoints.SourceModuleName()}.CalculateVolumeFromDiameter 7.5, 4.2, mFoo
End Sub

Public Sub WithMemberAccess()
    With {endpoints.SourceModuleName()}
        .CalculateVolumeFromDiameter 8.5, 4.2, mFoo
    End With
End Sub

Public Sub NonQualifiedAccess()
    CalculateVolumeFromDiameter 9.5, 4.2, mFoo
End Sub
";
            var referencingModuleName = "Module3";

            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints,
                    endpoints.ToSourceTuple(source),
                    endpoints.ToDestinationTuple(string.Empty),
                    (referencingModuleName, externalReferences, ComponentType.StandardModule));

            StringAssert.DoesNotContain("Public Sub CalculateVolumeFromDiameter(", refactoredCode.Source);
            StringAssert.Contains($"Public Sub CalculateVolumeFromDiameter(", refactoredCode.Destination);

            var module3Content = refactoredCode[referencingModuleName];

            StringAssert.Contains($"{endpoints.DestinationModuleName()}.CalculateVolumeFromDiameter 7.5, 4.2, mFoo", module3Content);
            StringAssert.Contains($"With {endpoints.SourceModuleName()}", module3Content);
            StringAssert.Contains($"{endpoints.DestinationModuleName()}.CalculateVolumeFromDiameter 8.5, 4.2, mFoo", module3Content);
            StringAssert.Contains($"{endpoints.DestinationModuleName()}.CalculateVolumeFromDiameter 9.5, 4.2, mFoo", module3Content);
        }

        [TestCase(MoveEndpoints.ClassToStd)]
        [TestCase(MoveEndpoints.FormToStd)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ExternallyReferencedMemberClassSource(MoveEndpoints endpoints)
        {
            var memberToMove = ("CalculateVolumeFromDiameter", DeclarationType.Procedure);
            var source =
$@"
Option Explicit

Public Sub CalculateVolumeFromDiameter(ByVal diameter As Single, ByVal height As Single, ByRef volume As Single)
    volume = height * 3.14 * (diameter / 2) ^ 2 
End Sub
";
            var instanceIdentifier = "mClassOrUserForm";
            var externalReferences =
$@"
Option Explicit

Private mFoo As Single

{ClassInstantiationBoilerPlate(instanceIdentifier, endpoints.SourceModuleName())}

Public Sub MemberAccess()
    {instanceIdentifier}.CalculateVolumeFromDiameter 7.5, 4.2, mFoo
End Sub

Public Sub WithMemberAccess()
    With {instanceIdentifier}
        .CalculateVolumeFromDiameter 8.5, 4.2, mFoo
    End With
End Sub
";
            ExecuteSingleTargetMoveThrowsExceptionTest(memberToMove, endpoints,
                endpoints.ToSourceTuple(source),
                endpoints.ToDestinationTuple(string.Empty),
                ("Module3", externalReferences, ComponentType.ClassModule));
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
            var memberToMove = ("RenameFile", DeclarationType.Procedure);
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
            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source);

            StringAssert.Contains($"{endpoints.DestinationModuleName()}.RenameFile mfileName, fileName", refactoredCode.Source);
            StringAssert.DoesNotContain($"{targetAccessibility} Sub RenameFile", refactoredCode.Source);

            StringAssert.Contains("Public Sub RenameFile", refactoredCode.Destination);
        }


        [TestCase(MoveEndpoints.StdToStd)]
        [TestCase(MoveEndpoints.ClassToStd)]
        [TestCase(MoveEndpoints.FormToStd)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void InternallyReferencedSubToStdModule(MoveEndpoints endpoints)
        {
            var memberToMove = ("CalculateVolume", DeclarationType.Procedure);

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
            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source);

            StringAssert.Contains($"{endpoints.DestinationModuleName()}.CalculateVolume(", refactoredCode.Source);
            StringAssert.DoesNotContain("Private Function CalculateVolume", refactoredCode.Source);

            StringAssert.Contains($"volume = area * height", refactoredCode.Destination);
            StringAssert.Contains($"Public Sub CalculateVolume(area", refactoredCode.Destination);
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
            var memberToMove = ("Foo", DeclarationType.Procedure);
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
            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source);

            StringAssert.AreEqualIgnoringCase(nameof(MoveMemberToStdModule), refactoredCode.StrategyName);

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
        public void ExclusiveCallChainNonExclusiveFieldMoveToStdModuleThrowsException(MoveEndpoints endpoints, string accessibility)
        {
            var memberToMove = ("Foo", DeclarationType.Procedure);
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
            ExecuteSingleTargetMoveThrowsExceptionTest(memberToMove, endpoints, source);
        }


        [TestCase(MoveEndpoints.StdToStd, "Public")]
        [TestCase(MoveEndpoints.StdToStd, "Private")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void UserSelectsAdditionalMethodToEnableMove(MoveEndpoints endpoints, string accessibility)
        {
            var memberToMove = ("Foo", DeclarationType.Procedure);
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
            MoveMemberModel modelAdjustment(MoveMemberModel model)
            {
                model.MoveableMemberSetByName("Goo").IsSelected = true;
                model.ChangeDestination(endpoints.DestinationModuleName(), ComponentType.StandardModule);
                return model;
            }
            var refactoredCode = RefactorTargets(memberToMove, endpoints, source, string.Empty, modelAdjustment);

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
        public void NonExclusiveCallChainPrivateMemberMoveToStdModuleStrategyNA(MoveEndpoints endpoints, string accessibility)
        {
            var memberToMove = ("Foo", DeclarationType.Procedure);
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
            ExecuteSingleTargetMoveThrowsExceptionTest(memberToMove, endpoints, source);
        }


        [TestCase(MoveEndpoints.StdToStd, "Public", nameof(MoveMemberToStdModule))]
        [TestCase(MoveEndpoints.ClassToStd, "Public", null)]
        [TestCase(MoveEndpoints.FormToStd, "Public", null)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void NonExclusiveCallChainPublicMember(MoveEndpoints endpoints, string accessibility, string expectedStrategy)
        {
            var memberToMove = ("Foo", DeclarationType.Procedure);
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
            if (expectedStrategy is null)
            {
                ExecuteSingleTargetMoveThrowsExceptionTest(memberToMove, endpoints, source);
                return;
            }

            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source);

            StringAssert.AreEqualIgnoringCase(expectedStrategy, refactoredCode.StrategyName);

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

        [TestCase(MoveEndpoints.StdToStd, "Public", nameof(MoveMemberToStdModule))]
        [TestCase(MoveEndpoints.StdToStd, "Private", nameof(MoveMemberToStdModule))]
        [TestCase(MoveEndpoints.ClassToStd, "Public", null)]
        [TestCase(MoveEndpoints.FormToStd, "Public", null)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ExclusiveCallChainMemberExternallyReferences(MoveEndpoints endpoints, string accessibility, string expectedStrategy)
        {
            var classInstanceCode = string.Empty;
            var memberToMove = ("Foo", DeclarationType.Procedure);
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
            var callSiteModuleName = "Module3";

            var memberAccessQualifier = endpoints.SourceModuleName();

            if (endpoints.IsClassSource() || endpoints.IsFormSource())
            {
                memberAccessQualifier = "classInstance";
                classInstanceCode = $"{ClassInstantiationBoilerPlate(memberAccessQualifier, endpoints.SourceModuleName())}";
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
            if (expectedStrategy is null)
            {
                ExecuteSingleTargetMoveThrowsExceptionTest(memberToMove, endpoints,
                    endpoints.ToSourceTuple(source),
                    endpoints.ToDestinationTuple(string.Empty),
                    (callSiteModuleName, callSiteCode, ComponentType.StandardModule));
                return;
            }

            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints,
                    endpoints.ToSourceTuple(source),
                    endpoints.ToDestinationTuple(string.Empty),
                    (callSiteModuleName, callSiteCode, ComponentType.StandardModule));

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
            StringAssert.Contains($"{endpoints.SourceModuleName()}.AddSix", destinationRefactored);
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
            var destinationModuleName = endpoints.DestinationModuleName();
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
            var refactoredCode = RefactorSingleTarget(("Foo", DeclarationType.Procedure), endpoints, source, destination);

            var destinationExpectedContent =
                @"
Option Explicit

Private Const LOG_IS_ENABLED = True

Public Entries As Long

Public Sub Foo(arg1 As Long)
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
                StringAssert.Contains(line, refactoredCode.Destination);
            }
        }

        [TestCase("Public", MoveEndpoints.StdToStd, nameof(MoveMemberToStdModule))]
        [TestCase("Private", MoveEndpoints.StdToStd, nameof(MoveMemberToStdModule))]
        [TestCase("Public", MoveEndpoints.ClassToStd, null)]
        [TestCase("Private", MoveEndpoints.ClassToStd, null)]
        [TestCase("Public", MoveEndpoints.FormToStd, null)]
        [TestCase("Private", MoveEndpoints.FormToStd, null)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ReferencesNonExclusiveProperty(string accessibility, MoveEndpoints endpoints, string expectedStrategy)
        {
            var memberToMove = ("Foo", DeclarationType.Procedure);
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
            if (expectedStrategy is null)
            {
                ExecuteSingleTargetMoveThrowsExceptionTest(memberToMove, endpoints, source);
                return;
            }

            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source);

            StringAssert.AreEqualIgnoringCase(expectedStrategy, refactoredCode.StrategyName);

            StringAssert.DoesNotContain("Sub Foo(arg1", refactoredCode.Source);
            StringAssert.Contains($"{accessibility} Sub Foo(arg1", refactoredCode.Destination);
            StringAssert.Contains($"arg1 = {endpoints.SourceModuleName()}", refactoredCode.Destination);
        }

        [TestCase(MoveEndpoints.StdToStd, nameof(MoveMemberToStdModule))]
        [TestCase(MoveEndpoints.ClassToStd, null)]
        [TestCase(MoveEndpoints.FormToStd, null)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void SupportMembersReferenceNonExclusiveBackingVariables(MoveEndpoints endpoints, string expectedStrategy)
        {
            var memberToMove = ("Foo", DeclarationType.Procedure);
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
            if (expectedStrategy is null)
            {
                ExecuteSingleTargetMoveThrowsExceptionTest(memberToMove, endpoints, source);
                return;
            }

            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source);

            StringAssert.AreEqualIgnoringCase(expectedStrategy, refactoredCode.StrategyName);

            StringAssert.Contains("Bar(", refactoredCode.Source);
            StringAssert.DoesNotContain("Sub Foo(arg1", refactoredCode.Source);
            StringAssert.Contains("Public Sub Foo(arg1", refactoredCode.Destination);
            StringAssert.Contains($"arg1 = {endpoints.SourceModuleName()}.Bar", refactoredCode.Destination);
            StringAssert.Contains($"{endpoints.SourceModuleName()}.Bar = arg1", refactoredCode.Destination);
            StringAssert.DoesNotContain("Bar(", refactoredCode.Destination);
        }


        [TestCase(MoveEndpoints.StdToStd, nameof(MoveMemberToStdModule))]
        [TestCase(MoveEndpoints.ClassToStd, null)]
        [TestCase(MoveEndpoints.FormToStd, null)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void SupportMembersReferenceNonExclusivePrivateMember(MoveEndpoints endpoints, string expectedStrategy)
        {
            var memberToMove = ("Foo", DeclarationType.Procedure);
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
            if (expectedStrategy is null)
            {
                ExecuteSingleTargetMoveThrowsExceptionTest(memberToMove, endpoints, source);
                return;
            }

            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source);

            StringAssert.AreEqualIgnoringCase(expectedStrategy, refactoredCode.StrategyName);

            StringAssert.Contains("Bar(", refactoredCode.Source);
            StringAssert.DoesNotContain("Sub Foo(arg1", refactoredCode.Source);
            StringAssert.Contains("Public Sub Foo(arg1", refactoredCode.Destination);
            StringAssert.Contains($"arg1 = {endpoints.SourceModuleName()}.Bar", refactoredCode.Destination);
            StringAssert.Contains($"{endpoints.SourceModuleName()}.Bar = arg1", refactoredCode.Destination);
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
            var memberToMove = ("Foo", DeclarationType.Procedure);
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
            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source);

            if (endpoints.IsStdModuleSource())
            {
                StringAssert.DoesNotContain("Sub Foo(arg1", refactoredCode.Source);
                StringAssert.Contains("Public Property Let Bar", refactoredCode.Source);
                StringAssert.Contains("Public Property Get Bar", refactoredCode.Source);

                StringAssert.Contains($"{accessibility} Sub Foo(arg1", refactoredCode.Destination);
                StringAssert.Contains($"arg1 = {endpoints.SourceModuleName()}.Bar * 10", refactoredCode.Destination);
                StringAssert.Contains($"{endpoints.SourceModuleName()}.Bar = arg1", refactoredCode.Destination);
            }
            else
            {
                StringAssert.DoesNotContain("Sub Foo(arg1", refactoredCode.Source);
                StringAssert.DoesNotContain("Public Property Let Bar", refactoredCode.Source);
                StringAssert.DoesNotContain("Public Property Get Bar", refactoredCode.Source);

                StringAssert.Contains($"{accessibility} Sub Foo(arg1", refactoredCode.Destination);
                StringAssert.Contains($"arg1 = Bar * 10", refactoredCode.Destination);
                StringAssert.Contains($" Bar = arg1", refactoredCode.Destination);
                StringAssert.Contains("Public Property Let Bar", refactoredCode.Destination);
                StringAssert.Contains("Public Property Get Bar", refactoredCode.Destination);
            }
        }

        [TestCase("Public", MoveEndpoints.StdToStd)]
        [TestCase("Private", MoveEndpoints.StdToStd)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ReferencesExternallyReferencedSupportProperty(string accessibility, MoveEndpoints endpoints)
        {
            var memberToMove = ("Foo", DeclarationType.Procedure);
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
            var externalReferencingCode =
$@"
Public Sub FooBar(arg1 As Long)
    {endpoints.SourceModuleName()}.Bar = arg1
End Sub
";
            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints,
                endpoints.ToSourceTuple(source),
                endpoints.ToDestinationTuple(string.Empty),
                ("Module3", externalReferencingCode, ComponentType.StandardModule));

            StringAssert.DoesNotContain("Sub Foo(arg1", refactoredCode.Source);

            StringAssert.Contains($"{accessibility} Sub Foo(arg1", refactoredCode.Destination);
            StringAssert.DoesNotContain("Public Property Let Bar", refactoredCode.Destination);
            StringAssert.DoesNotContain("Public Property Get Bar", refactoredCode.Destination);

            var module3Content = refactoredCode["Module3"];
            StringAssert.Contains($"{endpoints.SourceModuleName()}.Bar", refactoredCode.Destination);
        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void PlacesCodeInCorrectSpotRelativeToDeclareStmt()
        {
            var memberToMove = ("Fizz", DeclarationType.Procedure);
            var source = $@"
Option Explicit

Public Sub Fizz(arg1 As Long)
End Sub
";

            var destination = $@"
Option Explicit

Private mBar As Long

Declare Sub MessageBeep Lib ""User32"" (ByVal N As Long)

Public Sub DoesNothing()
End Sub
";
            var refactoredCode = RefactorSingleTarget(memberToMove, MoveEndpoints.StdToStd, source);

            StringAssert.DoesNotContain("Sub Foo(arg1", refactoredCode.Source);

            StringAssert.Contains($"Sub Fizz(arg1", refactoredCode.Destination);
            var idxOfFizz = refactoredCode.Destination.IndexOf("Sub Fizz");
            var idxOfDeclare = refactoredCode.Destination.IndexOf("Declare Sub");

            Assert.Less(idxOfDeclare, idxOfFizz);
        }


        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void DoesNotMoveMembersThatRaiseAnEvent()
        {
            var memberToMove = ("RaisesAnEvent", DeclarationType.Procedure);
            var endpoints = MoveEndpoints.ClassToStd;
            var source = $@"
Option Explicit

Public Event EventName(IDNumber As Long, ByRef Cancel As Boolean)

Public Sub RaisesAnEvent()
    RaiseEvent EventName(6, true)
End Sub";

            var eventSinkName = "CEventSink";
            var eventSinkContent = $@"
Option Explicit

Dim WithEvents TestEvents As {endpoints.SourceModuleName()}

Private Sub TestEvents_EventName(IDNumber As Long, ByRef Cancel As Boolean)
End Sub
";
            ExecuteSingleTargetMoveThrowsExceptionTest(memberToMove, endpoints,
                endpoints.ToSourceTuple(source),
                endpoints.ToDestinationTuple(string.Empty),
                (eventSinkName, eventSinkContent, ComponentType.ClassModule));
        }


        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void DoesNotMoveEventSink()
        {
            var eventSourceName = "CEventSource";

            var endpoints = MoveEndpoints.ClassToStd;

            var memberToMove = ("TestEvents_EventName", DeclarationType.Procedure);
            var source = $@"
Option Explicit

Dim WithEvents TestEvents As {eventSourceName}

Private Sub TestEvents_EventName(IDNumber As Long, ByRef Cancel As Boolean)
End Sub
";

            var eventSourceContent = $@"
Option Explicit

Public Event EventName(IDNumber As Long, ByRef Cancel As Boolean)

Public Sub RaisesAnEvent()
    RaiseEvent EventName(6, true)
End Sub";
            ExecuteSingleTargetMoveThrowsExceptionTest(memberToMove, endpoints,
                endpoints.ToSourceTuple(source),
                endpoints.ToDestinationTuple(string.Empty),
                (eventSourceName, eventSourceContent, ComponentType.ClassModule));
        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void DoesNotMoveInterfaceImplementationMember()
        {
            var memberToMove = ("ITestInterface_TestGet", DeclarationType.PropertyGet);
            var endpoints = MoveEndpoints.ClassToStd;
            var source = $@"
Option Explicit

Implements ITestInterface

Private mTestValue As Long

Private Property Get ITestInterface_TestGet() As Long
    ITestInterface_TestGet = mTestValue
End Property

";

            var interfaceDeclarationClass = "ITestInterface";
            var interfaceContent = $@"
Option Explicit

Public Property Get TestGet() As Long
End Property
";
            ExecuteSingleTargetMoveThrowsExceptionTest(memberToMove, endpoints,
                endpoints.ToSourceTuple(source),
                endpoints.ToDestinationTuple(string.Empty),
                (interfaceDeclarationClass, interfaceContent, ComponentType.ClassModule));
        }

    }
}
