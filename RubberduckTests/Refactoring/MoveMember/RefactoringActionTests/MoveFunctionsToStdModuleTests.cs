using NUnit.Framework;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.VBEditor.SafeComWrappers;
using System;

namespace RubberduckTests.Refactoring.MoveMember
{
    [TestFixture]
    public class MoveFunctionsToStdModuleTests : MoveMemberRefactoringActionTestSupportBase
    {
        [TestCase(MoveEndpoints.StdToClass)]
        [TestCase(MoveEndpoints.ClassToClass)]
        [TestCase(MoveEndpoints.FormToClass)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void SimpleMoveToClassModuleThrowsUnsupportedMoveException(MoveEndpoints endpoints)
        {
            var memberToMove = ("Foo", DeclarationType.Function);
            var source =
$@"
Option Explicit

Public Function Foo(arg1 As Long) As Long
    Foo = 10 * arg1
End Function
";
            ExecuteSingleTargetMoveThrowsExceptionTest(memberToMove, endpoints, source);
        }

        [TestCase(MoveEndpoints.StdToStd, "Public")]
        [TestCase(MoveEndpoints.StdToStd, "Private")]
        [TestCase(MoveEndpoints.ClassToStd, "Public")]
        [TestCase(MoveEndpoints.ClassToStd, "Private")]
        [TestCase(MoveEndpoints.FormToStd, "Public")]
        [TestCase(MoveEndpoints.FormToStd, "Private")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void SupportFunctionReferencesExclusiveSupportConstant(MoveEndpoints endpoints, string exclusiveFuncAccessibility)
        {
            var memberToMove = ("CalculateVolumeFromDiameter", DeclarationType.Function);
            var source =
$@"
Option Explicit

Private Const Pi As Single = 3.14

Public Function CalculateVolumeFromDiameter(ByVal diameter As Single, ByVal height As Single) As Single
    CalculateVolumeFromDiameter = CalculateVolume(diameter / 2, height)
End Function

{exclusiveFuncAccessibility} Function CalculateVolume(ByVal radius As Single, ByVal height As Single) As Single
    CalculateVolume = height * Pi * radius ^ 2
End Function
";

            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source);

            if (endpoints != MoveEndpoints.StdToStd)
            {
                var refactoredLines = refactoredCode.Destination.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var line in refactoredLines)
                {
                    //Moves everything from Source to Destination as-is
                    StringAssert.Contains(line, source);
                }
                return;
            }

            StringAssert.DoesNotContain("Public Function CalculateVolumeFromDiameter(", refactoredCode.Source);
            if (exclusiveFuncAccessibility.Equals(Tokens.Private))
            {
                StringAssert.DoesNotContain($"{exclusiveFuncAccessibility} Function CalculateVolume(", refactoredCode.Source);
                StringAssert.DoesNotContain($"Pi As Single", refactoredCode.Source);
            }
            else
            {
                StringAssert.Contains($"{exclusiveFuncAccessibility} Function CalculateVolume(", refactoredCode.Source);
                StringAssert.Contains($"Pi As Single", refactoredCode.Source);
            }

            if (exclusiveFuncAccessibility.Equals(Tokens.Private))
            {
                StringAssert.Contains($"Private Const Pi As Single", refactoredCode.Destination);
                StringAssert.Contains($"{exclusiveFuncAccessibility} Function CalculateVolume(", refactoredCode.Destination);
                StringAssert.Contains($"CalculateVolumeFromDiameter = CalculateVolume(", refactoredCode.Destination);
            }
            else
            {
                StringAssert.DoesNotContain($"Private Const Pi As Single", refactoredCode.Destination);
                StringAssert.DoesNotContain($"{exclusiveFuncAccessibility} Function CalculateVolume(", refactoredCode.Destination);
                StringAssert.Contains($"CalculateVolumeFromDiameter = {endpoints.SourceModuleName()}.CalculateVolume(", refactoredCode.Destination);
            }
            StringAssert.Contains("Public Function CalculateVolumeFromDiameter(", refactoredCode.Destination);
        }

        [TestCase(MoveEndpoints.StdToStd, "Public")]
        [TestCase(MoveEndpoints.StdToStd, "Private")]
        [TestCase(MoveEndpoints.ClassToStd, "Public")]
        [TestCase(MoveEndpoints.ClassToStd, "Private")]
        [TestCase(MoveEndpoints.FormToStd, "Public")]
        [TestCase(MoveEndpoints.FormToStd, "Private")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MovedFunctionReferencesExclusiveSupportConstantSelectAllMembers(MoveEndpoints endpoints, string exclusiveFuncAccessibility)
        {
            var memberToMove = ("CalculateVolumeFromDiameter", DeclarationType.Function);
            var source =
$@"
Option Explicit

Private Const Pi As Single = 3.14

Public Function CalculateVolumeFromDiameter(ByVal diameter As Single, ByVal height As Single) As Single
    CalculateVolumeFromDiameter = CalculateVolume(diameter / 2, height)
End Function

{exclusiveFuncAccessibility} Function CalculateVolume(ByVal radius As Single, ByVal height As Single) As Single
    CalculateVolume = height * Pi * radius ^ 2
End Function
";

            Func<MoveMemberModel, MoveMemberModel> modelAdjustment = model =>
            {
                model.MoveableMemberSetByName("CalculateVolume").IsSelected = true;
                model.ChangeDestination(endpoints.DestinationModuleName(), ComponentType.StandardModule);
                return model;
            };

            var refactoredCode = RefactorTargets(memberToMove, endpoints, source, string.Empty, modelAdjustment);

            StringAssert.DoesNotContain("Public Function CalculateVolumeFromDiameter(", refactoredCode.Source);
            StringAssert.DoesNotContain($"{exclusiveFuncAccessibility} Function CalculateVolume(", refactoredCode.Source);
            StringAssert.DoesNotContain($"Pi As Single", refactoredCode.Source);

            StringAssert.Contains($"Private Const Pi As Single", refactoredCode.Destination);
            StringAssert.Contains("Public Function CalculateVolumeFromDiameter(", refactoredCode.Destination);
            StringAssert.Contains("CalculateVolumeFromDiameter = CalculateVolume(", refactoredCode.Destination);
            StringAssert.Contains($"{exclusiveFuncAccessibility} Function CalculateVolume(", refactoredCode.Destination);
        }

        [TestCase(MoveEndpoints.StdToStd, "Public Const Pi As Single = 3.14", false)]
        [TestCase(MoveEndpoints.StdToStd, "Public Pi As Single", false)]
        [TestCase(MoveEndpoints.ClassToStd, "Public Pi As Single", true)]
        [TestCase(MoveEndpoints.FormToStd, "Public Pi As Single", true)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MovedFunctionReferencesNonExclusivePublicSupportNonMember(MoveEndpoints endpoints, string nonMemberDeclaration, bool throwsException)
        {
            var memberToMove = ("CalculateVolumeFromDiameter", DeclarationType.Function);
            var source =
$@"
Option Explicit

{nonMemberDeclaration}

Public Function CalculateVolumeFromDiameter(ByVal diameter As Single, ByVal height As Single) As Single
    CalculateVolumeFromDiameter = height * Pi * (diameter / 2) ^ 2
End Function

Public Function CalculateCircumferenceFromDiameter(ByVal diameter As Single) As Single
    CalculateCircumferenceFromDiameter = diameter * Pi
End Function";

            if (throwsException)
            {
                ExecuteSingleTargetMoveThrowsExceptionTest(memberToMove, endpoints, source);
                return;
            }

            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source);

            StringAssert.DoesNotContain("Public Function CalculateVolumeFromDiameter(", refactoredCode.Source);
            StringAssert.Contains($"Public Function CalculateCircumferenceFromDiameter(", refactoredCode.Source);

            StringAssert.Contains($"Public Function CalculateVolumeFromDiameter(", refactoredCode.Destination);
            StringAssert.Contains($"height * {endpoints.SourceModuleName()}.Pi", refactoredCode.Destination);
        }

        [TestCase("Private Const Pi As Single = 3.14")]
        [TestCase("Private Pi As Single")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MovedFunctionReferencesNonExclusivePrivateSupportNonMember(string nonMemberDeclaration)
        {
            var memberToMove = ("CalculateVolumeFromDiameter", DeclarationType.Function);
            var source =
$@"
Option Explicit

{nonMemberDeclaration}

Public Function CalculateVolumeFromDiameter(ByVal diameter As Single, ByVal height As Single) As Single
    CalculateVolumeFromDiameter = height * Pi * (diameter / 2) ^ 2
End Function

Public Function CalculateCircumferenceFromDiameter(ByVal diameter As Single) As Single
    CalculateCircumferenceFromDiameter = diameter * Pi
End Function";

            ExecuteSingleTargetMoveThrowsExceptionTest(memberToMove, MoveEndpoints.StdToStd, source);
        }

        [TestCase(MoveEndpoints.StdToStd, "Public", false)]
        [TestCase(MoveEndpoints.StdToStd, "Private", true)]
        [TestCase(MoveEndpoints.ClassToStd, "Public", true)]
        [TestCase(MoveEndpoints.ClassToStd, "Private", true)]
        [TestCase(MoveEndpoints.FormToStd, "Public", true)]
        [TestCase(MoveEndpoints.FormToStd, "Private", true)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ReferencesNonExclusiveMember(MoveEndpoints endpoints, string exclusiveFuncAccessibility, bool throwsException)
        {
            var memberToMove = ("CalculateVolumeFromDiameter", DeclarationType.Function);
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
            if (throwsException)
            {
                ExecuteSingleTargetMoveThrowsExceptionTest(memberToMove, endpoints, source);
                return;
            }

            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source);

            StringAssert.DoesNotContain("Public Function CalculateVolumeFromDiameter(", refactoredCode.Source);
            StringAssert.Contains($"{exclusiveFuncAccessibility} Function CalculateVolume(", refactoredCode.Source);

            StringAssert.Contains($"Public Function CalculateVolumeFromDiameter(", refactoredCode.Destination);
            StringAssert.Contains($"CalculateVolumeFromDiameter = {endpoints.SourceModuleName()}.CalculateVolume(", refactoredCode.Destination);
            StringAssert.DoesNotContain($"{exclusiveFuncAccessibility} Function CalculateVolume(", refactoredCode.Destination);
        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ExternallyReferencedMemberStdToStd()
        {
            var memberToMove = ("CalculateVolumeFromDiameter", DeclarationType.Function);
            var endpoints = MoveEndpoints.StdToStd;
            var source =
$@"
Option Explicit

Public Function CalculateVolumeFromDiameter(ByVal diameter As Single, ByVal height As Single) As Single
    CalculateVolumeFromDiameter = height * 3.14 * (diameter / 2) ^ 2 
End Function
";
            var externalReferences =
$@"
Option Explicit

Private mFoo As Single

Public Sub MemberAccess()
    mFoo = {endpoints.SourceModuleName()}.CalculateVolumeFromDiameter(7.5, 4.2)
End Sub

Public Sub WithMemberAccess()
    With {endpoints.SourceModuleName()}
        mFoo = .CalculateVolumeFromDiameter(8.5, 4.2)
    End With
End Sub

Public Sub NonQualifiedAccess()
    mFoo = CalculateVolumeFromDiameter(9.5, 4.2)
End Sub
";
            var referencingModuleName = "Module3";

            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints,
                                                    endpoints.ToSourceTuple(source),
                                                    endpoints.ToDestinationTuple(string.Empty),
                                                    (referencingModuleName, externalReferences, ComponentType.StandardModule));


            StringAssert.DoesNotContain("Public Function CalculateVolumeFromDiameter(", refactoredCode.Source);
            StringAssert.Contains($"Public Function CalculateVolumeFromDiameter(", refactoredCode.Destination);

            var module3Content = refactoredCode[referencingModuleName];

            StringAssert.Contains($"{endpoints.DestinationModuleName()}.CalculateVolumeFromDiameter(7.5, 4.2)", module3Content);
            StringAssert.Contains($"With {endpoints.SourceModuleName()}", module3Content);
            StringAssert.Contains($"{endpoints.DestinationModuleName()}.CalculateVolumeFromDiameter(8.5, 4.2)", module3Content);
            StringAssert.Contains($"{endpoints.DestinationModuleName()}.CalculateVolumeFromDiameter(9.5, 4.2)", module3Content);
        }

        [TestCase(MoveEndpoints.ClassToStd)]
        [TestCase(MoveEndpoints.FormToStd)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ExternallyReferencedMemberClassSourceThrowsException(MoveEndpoints endpoints)
        {
            var memberToMove = ("CalculateVolumeFromDiameter", DeclarationType.Function);
            var source =
$@"
Option Explicit

Public Function CalculateVolumeFromDiameter(ByVal diameter As Single, ByVal height As Single) As Single
    CalculateVolumeFromDiameter = height * 3.14 * (diameter / 2) ^ 2 
End Function
";
            var instanceIdentifier = "mClassOrUserForm";
            var externalReferences =
$@"
Option Explicit

Private mFoo As Single

{ClassInstantiationBoilerPlate(instanceIdentifier, endpoints.SourceModuleName())}

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

            ExecuteSingleTargetMoveThrowsExceptionTest(memberToMove, endpoints,
                                                    endpoints.ToSourceTuple(source),
                                                    endpoints.ToDestinationTuple(string.Empty),
                                                    (referencingModuleName, externalReferences, ComponentType.StandardModule));
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
            var memberToMove = ("TryRenameFile", DeclarationType.Function);
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
            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source);

            StringAssert.Contains($"{endpoints.DestinationModuleName()}.TryRenameFile(mfileName, fileName)", refactoredCode.Source);
            StringAssert.DoesNotContain($"{targetAccessibility} Function TryRenameFile", refactoredCode.Source);

            StringAssert.Contains("Public Function TryRenameFile", refactoredCode.Destination);
        }

        [TestCase(MoveEndpoints.StdToStd)]
        [TestCase(MoveEndpoints.ClassToStd)]
        [TestCase(MoveEndpoints.FormToStd)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void InternallyReferencedFunctionToStdModule(MoveEndpoints endpoints)
        {
            var memberToMove = ("CalculateVolume", DeclarationType.Function);

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
            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source);

            StringAssert.Contains($"{endpoints.DestinationModuleName()}.CalculateVolume(", refactoredCode.Source);
            StringAssert.DoesNotContain("Private Function CalculateVolume", refactoredCode.Source);

            StringAssert.Contains($"CalculateVolume = area * height", refactoredCode.Destination);
            StringAssert.Contains($"Public Function CalculateVolume(area", refactoredCode.Destination);
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
            var memberToMove = ("Foo", DeclarationType.Function);
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
            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source);

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

        [TestCase(MoveEndpoints.StdToStd, "Private")]
        [TestCase(MoveEndpoints.StdToStd, "Public")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ExclusiveCallChainNonExclusiveFieldAccessedViaPublicSupportMember(MoveEndpoints endpoints, string accessibility)
        {
            var memberToMove = ("Foo", DeclarationType.Function);
            var source =
$@"
Option Explicit

Private mfoo As Long

{accessibility} Function Foo(arg1 As Long) As Long
    Foo = Bar(arg1)
End Function

Public Sub Goo(arg1 As Long)
    mfoo = arg1
End Sub

Public Function Bar(arg1 As Long) As Long
    Bar = Barn(arg1) + 4
End Function

Private Function Barn(arg1 As Long) As Long
    Barn = arg1 + mfoo
End Function
";
            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source);

            var sourceRefactored = refactoredCode.Source;
            StringAssert.Contains("Private mfoo As Long", sourceRefactored);
            StringAssert.Contains("Function Bar(", sourceRefactored);
            StringAssert.Contains("Function Barn(", sourceRefactored);

            var destinationRefactored = refactoredCode.Destination;
            StringAssert.Contains($"{accessibility} Function Foo", destinationRefactored);
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
            var memberToMove = ("Foo", DeclarationType.Function);
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
            ExecuteSingleTargetMoveThrowsExceptionTest(memberToMove, endpoints, source);
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
            var memberToMove = ("Foo", DeclarationType.Function);
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

            ExecuteSingleTargetMoveThrowsExceptionTest(memberToMove, endpoints, source);
        }

        [TestCase(MoveEndpoints.StdToStd, "Public", false)]
        [TestCase(MoveEndpoints.ClassToStd, "Public", true)]
        [TestCase(MoveEndpoints.FormToStd, "Public", true)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void NonExclusiveCallChainPublicMember(MoveEndpoints endpoints, string accessibility, bool throwsException)
        {
            var memberToMove = ("Foo", DeclarationType.Function);
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

            if (throwsException)
            {
                ExecuteSingleTargetMoveThrowsExceptionTest(memberToMove, endpoints, source);
                return;
            }

            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source);

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

        [TestCase(MoveEndpoints.StdToStd, "Public")]
        [TestCase(MoveEndpoints.StdToStd, "Private")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ExclusiveCallChainMemberExternallyReferencedStdModuleSource(MoveEndpoints endpoints, string accessibility)
        {
            var classInstanceCode = string.Empty;
            var memberToMove = ("Foo", DeclarationType.Function);
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

        [TestCase(MoveEndpoints.ClassToStd, "Public")]
        [TestCase(MoveEndpoints.FormToStd, "Public")]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ExclusiveCallChainMemberExternallyReferencedClassSource(MoveEndpoints endpoints, string accessibility)
        {
            var classInstanceCode = string.Empty;
            var memberToMove = ("Foo", DeclarationType.Function);
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
";
            ExecuteSingleTargetMoveThrowsExceptionTest(memberToMove, endpoints,
                    endpoints.ToSourceTuple(source),
                    endpoints.ToDestinationTuple(string.Empty),
                    (callSiteModuleName, callSiteCode, ComponentType.StandardModule));
        }

        [TestCase(MoveEndpoints.StdToStd)]
        [TestCase(MoveEndpoints.ClassToStd)]
        [TestCase(MoveEndpoints.FormToStd)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void RemovesMemberAccessInDestination(MoveEndpoints endpoints)
        {
            var destinationModuleName = MoveEndpoints.StdToStd.DestinationModuleName();
            var memberToMove = ("Foo", DeclarationType.Function);
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
            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source, destination);

            var destinationExpectedContent =
                @"
Option Explicit

Private Const LOG_IS_ENABLED = True

Public Entries As Long

Public Function Foo(arg1 As Long)
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
                StringAssert.Contains(line, refactoredCode.Destination);
            }
        }


        [TestCase(MoveEndpoints.StdToStd)]
        [TestCase(MoveEndpoints.ClassToStd)]
        [TestCase(MoveEndpoints.FormToStd)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void RemovesWithMemberAccessInDestination(MoveEndpoints endpoints)
        {
            var memberToMove = ("Foo", DeclarationType.Function);
            var destinationModuleName = MoveEndpoints.StdToStd.DestinationModuleName();
            var source =
$@"
Option Explicit

Private mfoo As Long
Private mgoo As Long

Public Function Foo(arg1 As Long) As Long
    mfoo = arg1 * 10
    With {destinationModuleName}
        If .LogIsEnabled Then
            .Log ""Foo called""
            .Entries = .Entries + 1
        Endif
    End With
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
            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source, destination);

            var destinationExpectedContent =
                @"
Option Explicit

Private Const LOG_IS_ENABLED = True

Public Entries As Long

Public Function Foo(arg1 As Long)
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
                StringAssert.Contains(line, refactoredCode.Destination);
            }
        }

        [TestCase("Public", MoveEndpoints.StdToStd, false)]
        [TestCase("Private", MoveEndpoints.StdToStd, false)]
        [TestCase("Public", MoveEndpoints.ClassToStd, true)]
        [TestCase("Private", MoveEndpoints.ClassToStd, true)]
        [TestCase("Public", MoveEndpoints.FormToStd, true)]
        [TestCase("Private", MoveEndpoints.FormToStd, true)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ReferencesNonExclusiveProperty(string accessibility, MoveEndpoints endpoints, bool throwsException)
        {
            var sourceModuleName = endpoints.SourceModuleName();
            var destinationModuleName = endpoints.DestinationModuleName();

            var memberToMove = ("Foo", DeclarationType.Function);
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

            if (throwsException)
            {
                ExecuteSingleTargetMoveThrowsExceptionTest(memberToMove, endpoints, source);
                return;
            }

            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source);

            StringAssert.DoesNotContain("Function Foo(arg1", refactoredCode.Source);
            StringAssert.Contains($"{accessibility} Function Foo(arg1", refactoredCode.Destination);
            StringAssert.Contains($"arg1 = {sourceModuleName}.Bar * 10", refactoredCode.Destination);
        }

        [TestCase(MoveEndpoints.StdToStd, false)]
        [TestCase(MoveEndpoints.ClassToStd, true)]
        [TestCase(MoveEndpoints.FormToStd, true)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void SupportMemberReferencesNonExclusiveBackingVariable( MoveEndpoints endpoints, bool throwsException)
        {
            var memberToMove = ("Foo", DeclarationType.Function);
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

            if (throwsException)
            {
                ExecuteSingleTargetMoveThrowsExceptionTest(memberToMove, endpoints, source);
                return;
            }

            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source);

            StringAssert.Contains("Bar(", refactoredCode.Source);
            StringAssert.DoesNotContain("Function Foo(arg1", refactoredCode.Source);
            StringAssert.Contains("Public Function Foo(arg1", refactoredCode.Destination);
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
            var memberToMove = ("Foo", DeclarationType.Function);
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
            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints, source);

            if (!endpoints.IsStdModuleSource())
            {
                StringAssert.DoesNotContain("Function Foo(arg1", refactoredCode.Source);
                StringAssert.DoesNotContain("Public Property Let Bar", refactoredCode.Source);
                StringAssert.DoesNotContain("Public Property Get Bar", refactoredCode.Source);
                StringAssert.DoesNotContain($"arg1 = Bar * 10", refactoredCode.Source);
                StringAssert.DoesNotContain($" Bar = arg1", refactoredCode.Source);

                StringAssert.Contains($"{accessibility} Function Foo(arg1", refactoredCode.Destination);
                StringAssert.Contains("Public Property Let Bar", refactoredCode.Destination);
                StringAssert.Contains("Public Property Get Bar", refactoredCode.Destination);
                StringAssert.Contains($"arg1 = Bar * 10", refactoredCode.Destination);
                StringAssert.Contains($" Bar = arg1", refactoredCode.Destination);
            }
            else
            {
                StringAssert.DoesNotContain("Function Foo(arg1", refactoredCode.Source);
                StringAssert.Contains("Public Property Let Bar", refactoredCode.Source);
                StringAssert.Contains("Public Property Get Bar", refactoredCode.Source);
                StringAssert.DoesNotContain($"arg1 = Bar * 10", refactoredCode.Source);
                StringAssert.DoesNotContain($" Bar = arg1", refactoredCode.Source);

                StringAssert.Contains($"{accessibility} Function Foo(arg1", refactoredCode.Destination);
                StringAssert.Contains($"arg1 = {endpoints.SourceModuleName()}.Bar * 10", refactoredCode.Destination);
                StringAssert.Contains($"{endpoints.SourceModuleName()}.Bar = arg1", refactoredCode.Destination);
            }
        }

        [TestCase("Public", MoveEndpoints.StdToStd)]
        [TestCase("Private", MoveEndpoints.StdToStd)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ReferencesExternallyReferencedSupportProperty(string accessibility, MoveEndpoints endpoints)
        {
            var memberToMove = ("Foo", DeclarationType.Function);
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

            StringAssert.DoesNotContain("Function Foo(arg1", refactoredCode.Source);

            StringAssert.Contains($"{accessibility} Function Foo(arg1", refactoredCode.Destination);
            StringAssert.DoesNotContain("Public Property Let Bar", refactoredCode.Destination);
            StringAssert.DoesNotContain("Public Property Get Bar", refactoredCode.Destination);

            var module3Content = refactoredCode["Module3"];
            StringAssert.Contains($"{endpoints.SourceModuleName()}.Bar", refactoredCode.Destination);

        }

        [Test]
        [Category(nameof(NameConflictFinder))]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void CorrectsMemberNameCollisionInDestination()
        {
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

Public multiplier As Long

Public Function Goo(arg1 As Long) 
    Goo = multiplier * arg1
End Function
";

            var refactoredCode = RefactorSingleTarget(memberToMove, MoveEndpoints.StdToStd, source, destination);

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
                StringAssert.Contains(line, refactoredCode.Destination);
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void SetsNewMemberNameAtExternalReferences()
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
    mBar = {endpoints.SourceModuleName()}.Goo
End Sub

Public Sub WithMemberAccess(arg2 As Long)
    With {endpoints.SourceModuleName()}
        mBar = .Goo
    End With
End Sub

Public Sub NonQualified(arg3 As Long)
    mBar = Goo + arg3
End Sub
";
            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints,
                         endpoints.ToSourceTuple(source),
                         endpoints.ToDestinationTuple(destination),
                         (callSiteModuleName, callSiteCode, ComponentType.StandardModule));

            var callSiteExpectedContent =
$@"
Option Explicit

Private mBar As Long

Public Sub MemberAccess(arg1 As Long)
    mBar = {endpoints.DestinationModuleName()}.Goo1
End Sub

Public Sub WithMemberAccess(arg2 As Long)
    With {endpoints.SourceModuleName()}
        mBar = {endpoints.DestinationModuleName()}.Goo1
    End With
End Sub

Public Sub NonQualified(arg3 As Long)
    mBar = {endpoints.DestinationModuleName()}.Goo1 + arg3
End Sub
";
            var expectedLines = callSiteExpectedContent.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var line in expectedLines)
            {
                StringAssert.Contains(line, refactoredCode[callSiteModuleName]);
            }
        }

        [TestCase(MoveEndpoints.StdToStd)]
        [TestCase(MoveEndpoints.ClassToStd)]
        [TestCase(MoveEndpoints.FormToStd)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ContainsObjectField(MoveEndpoints endpoints)
        {
            var memberToMove = ("FooMath", DeclarationType.Function);
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
            var refactoredCode = RefactorSingleTarget(memberToMove, endpoints,
                    endpoints.ToSourceTuple(source),
                    endpoints.ToDestinationTuple(string.Empty),
                    ("ObjectClass", objectClass, ComponentType.ClassModule));

            var sourceRefactored = refactoredCode.Source;
            StringAssert.DoesNotContain("FooMath", sourceRefactored);
            StringAssert.DoesNotContain("mObj", sourceRefactored);

            var destinationRefactored = refactoredCode.Destination;
            StringAssert.Contains("FooMath", destinationRefactored);
            StringAssert.Contains("mObj", destinationRefactored);
        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MovesPrivateMethodExclusiveToSelectedMember()
        {
            var memberToMove = ("Foo", DeclarationType.Function);
            var source = $@"
Option Explicit

Private mBar As Long

Public Sub InitializeModule(arg1 As Long)
    mBar = arg1
End Sub

Public Function Foo(arg1 As Long) As Long
    arg1 = Bar * 10
    Bar = arg1
    LogFoo
    Foo = arg1
End Function

Public Property Let Bar(arg1 As Long)
    LogBar
    mBar = arg1
End Property

Public Property Get Bar() As Long
    LogBar
    Bar = mBar
End Property

'LogFoo is a direct dependency of Foo...it has to move
Private Sub LogFoo()
End Sub

Private Sub LogBar()
End Sub
";
            var refactoredCode = RefactorSingleTarget(memberToMove, MoveEndpoints.StdToStd, source);

            StringAssert.Contains("Bar(", refactoredCode.Source);
            StringAssert.Contains("LogBar(", refactoredCode.Source);

            StringAssert.Contains("Public Function Foo(arg1", refactoredCode.Destination);
            StringAssert.Contains($"Private Sub LogFoo", refactoredCode.Destination);
            StringAssert.DoesNotContain("LogBar(", refactoredCode.Destination);
            StringAssert.Contains($"{MoveEndpoints.StdToStd.SourceModuleName()}.Bar = arg1", refactoredCode.Destination);
            StringAssert.DoesNotContain("Bar(", refactoredCode.Destination);
        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void PrivateMethodForcesPublicSupportMemberToMove()
        {
            var memberToMove = ("Foo", DeclarationType.Function);
            var source = $@"
Option Explicit

Private mBar As Long

Public Function Foo(arg1 As Long) As Long
    arg1 = Bar * 10
    Bar = arg1
    LogFoo
    Foo = arg1
End Function

Public Property Let Bar(arg1 As Long)
    LogFoo
    mBar = arg1
End Property

Public Property Get Bar() As Long
    LogFoo
    Bar = mBar
End Property

Private Sub LogFoo()
End Sub
";
            var refactoredCode = RefactorSingleTarget(memberToMove, MoveEndpoints.StdToStd, source);

            StringAssert.DoesNotContain("Bar(", refactoredCode.Source);
            StringAssert.DoesNotContain("LogFoo(", refactoredCode.Source);

            StringAssert.Contains("Public Function Foo(arg1", refactoredCode.Destination);
            StringAssert.Contains($"Private Sub LogFoo", refactoredCode.Destination);
            StringAssert.Contains("Bar(", refactoredCode.Destination);
            StringAssert.Contains("mBar As Long", refactoredCode.Destination);
        }
    }
}
