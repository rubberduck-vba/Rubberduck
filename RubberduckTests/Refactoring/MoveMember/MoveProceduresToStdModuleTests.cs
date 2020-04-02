using NUnit.Framework;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.VBEditor.SafeComWrappers;
using System;
using Support = RubberduckTests.Refactoring.MoveMember.MoveMemberTestSupport;


namespace RubberduckTests.Refactoring.MoveMember
{
    [TestFixture]
    public class MoveProceduresToStdModuleTests : MoveMemberRefactoringActionTestSupportBase
    {
        private const string ThisStrategy = nameof(MoveMemberToStdModule);
        private const DeclarationType ThisDeclarationType = DeclarationType.Procedure;

        [TestCase(MoveEndpoints.StdToClass)]
        [TestCase(MoveEndpoints.ClassToClass)]
        [TestCase(MoveEndpoints.FormToClass)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void SimpleMoveToClassModule_NoStrategy(MoveEndpoints endpoints)
        {
            var source = $@"
Option Explicit

Public Sub Log()
End Sub";

            var moveDefinition = new TestMoveDefinition(endpoints, ("Log", ThisDeclarationType), sourceContent: source);

            
            var refactoredCode = ExecuteTest(moveDefinition);

            Assert.AreEqual(null, refactoredCode.StrategyName);
        }

        [TestCase(MoveEndpoints.StdToStd, "Public", ThisStrategy)]
        [TestCase(MoveEndpoints.StdToStd, "Private", ThisStrategy)]
        [TestCase(MoveEndpoints.ClassToStd, "Public", ThisStrategy)]
        [TestCase(MoveEndpoints.ClassToStd, "Private", ThisStrategy)]
        [TestCase(MoveEndpoints.FormToStd, "Public", ThisStrategy)]
        [TestCase(MoveEndpoints.FormToStd, "Private", ThisStrategy)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void SupportSubReferencesExclusiveSupportConstant(MoveEndpoints endpoints, string exclusiveFuncAccessibility, string expectedStrategy)
        {
            var memberToMove = "CalculateVolumeFromDiameter";
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

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType), sourceContent: source);

            
            var refactoredCode = ExecuteTest(moveDefinition);

            StringAssert.AreEqualIgnoringCase(expectedStrategy, refactoredCode.StrategyName);

            if (moveDefinition.IsStdModuleSource)
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
                    StringAssert.Contains($"{moveDefinition.SourceModuleName}.CalculateVolume(diameter / 2, height, volume)", refactoredCode.Destination);
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

        [TestCase(MoveEndpoints.StdToStd, "Public Const Pi As Single = 3.14", ThisStrategy)]
        [TestCase(MoveEndpoints.StdToStd, "Public Pi As Single", ThisStrategy)]
        [TestCase(MoveEndpoints.ClassToStd, "Public Pi As Single", null)]
        [TestCase(MoveEndpoints.FormToStd, "Public Pi As Single", null)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MovedSubReferencesNonExclusivePublicSupportNonMember(MoveEndpoints endpoints, string nonMemberDeclaration, string expectedStrategy)
        {
            var memberToMove = "CalculateVolumeFromDiameter";
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

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType), sourceContent: source);

            
            var refactoredCode = ExecuteTest(moveDefinition);

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

        [TestCase(MoveEndpoints.StdToStd, "Private Const Pi As Single = 3.14", null)]
        [TestCase(MoveEndpoints.StdToStd, "Private Pi As Single", null)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void ReferencesNonExclusivePrivateSupportNonMember(MoveEndpoints endpoints, string nonMemberDeclaration, string expectedStrategy)
        {
            var memberToMove = "CalculateVolumeFromDiameter";
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

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType), sourceContent: source);

            
            var refactoredCode = ExecuteTest(moveDefinition);

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

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType), sourceContent: source);
            
            var refactoredCode = ExecuteTest(moveDefinition);

            if (expectedStrategyName is null)
            {
                return;
            }

            StringAssert.DoesNotContain("Public Sub Foo(arg1 As Long)", refactoredCode.Source);

            StringAssert.Contains("Public Sub Foo(arg1 As Long)", refactoredCode.Destination);
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


            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, (memberToMove, ThisDeclarationType), sourceContent: source);

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

            
            var refactoredCode = ExecuteTest(moveDefinition);

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

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType), sourceContent: source);

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

            
            var refactoredCode = ExecuteTest(moveDefinition);

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

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType), sourceContent: source);

            
            var refactoredCode = ExecuteTest(moveDefinition);

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

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType), sourceContent: source);

            
            var refactoredCode = ExecuteTest(moveDefinition);

            StringAssert.AreEqualIgnoringCase(ThisStrategy, refactoredCode.StrategyName);
            StringAssert.Contains($"{moveDefinition.DestinationModuleName}.CalculateVolume(", refactoredCode.Source);
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

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType), sourceContent: source);

            
            var refactoredCode = ExecuteTest(moveDefinition);

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

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType), sourceContent: source);

            
            var refactoredCode = ExecuteTest(moveDefinition);

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

            moveDefinition.SetEndpointContent(source);
            moveDefinition.AddSelectedDeclaration("Goo", DeclarationType.Procedure);
            var refactoredCode = ExecuteTest(moveDefinition);

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

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType), sourceContent: source);

            
            var refactoredCode = ExecuteTest(moveDefinition);

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

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType), sourceContent: source);

            
            var refactoredCode = ExecuteTest(moveDefinition);

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

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType), sourceContent: source);

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

            
            var refactoredCode = ExecuteTest(moveDefinition);

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

            moveDefinition.SetEndpointContent(source, destination);
            var refactorResults = ExecuteTest(moveDefinition);

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

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType), sourceContent: source);
            
            var refactoredCode = ExecuteTest(moveDefinition);

            StringAssert.AreEqualIgnoringCase(expectedStrategy, refactoredCode.StrategyName);
            if (expectedStrategy is null)
            {
                return;
            }

            StringAssert.DoesNotContain("Sub Foo(arg1", refactoredCode.Source);
            StringAssert.Contains($"{accessibility} Sub Foo(arg1", refactoredCode.Destination);
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

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType), sourceContent: source);
            
            var refactoredCode = ExecuteTest(moveDefinition);

            StringAssert.AreEqualIgnoringCase(expectedStrategy, refactoredCode.StrategyName);
            if (expectedStrategy is null)
            {
                return;
            }

            StringAssert.Contains("Bar(", refactoredCode.Source);
            StringAssert.DoesNotContain("Sub Foo(arg1", refactoredCode.Source);
            StringAssert.Contains("Public Sub Foo(arg1", refactoredCode.Destination);
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

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType), sourceContent: source);
            
            var refactoredCode = ExecuteTest(moveDefinition);

            StringAssert.AreEqualIgnoringCase(expectedStrategy, refactoredCode.StrategyName);
            if (expectedStrategy is null)
            {
                return;
            }

            StringAssert.Contains("Bar(", refactoredCode.Source);
            StringAssert.DoesNotContain("Sub Foo(arg1", refactoredCode.Source);
            StringAssert.Contains("Public Sub Foo(arg1", refactoredCode.Destination);
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

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType), sourceContent: source);
            
            var refactoredCode = ExecuteTest(moveDefinition);

            StringAssert.AreEqualIgnoringCase(ThisStrategy, refactoredCode.StrategyName);

            if (moveDefinition.IsStdModuleSource)
            {
                StringAssert.DoesNotContain("Sub Foo(arg1", refactoredCode.Source);
                StringAssert.Contains("Public Property Let Bar", refactoredCode.Source);
                StringAssert.Contains("Public Property Get Bar", refactoredCode.Source);

                StringAssert.Contains($"{accessibility} Sub Foo(arg1", refactoredCode.Destination);
                StringAssert.Contains($"arg1 = {moveDefinition.SourceModuleName}.Bar * 10", refactoredCode.Destination);
                StringAssert.Contains($"{moveDefinition.SourceModuleName}.Bar = arg1", refactoredCode.Destination);
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

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, ThisDeclarationType), sourceContent: source);

            var externalReferencingCode =
$@"
Public Sub FooBar(arg1 As Long)
    {moveDefinition.SourceModuleName}.Bar = arg1
End Sub
";
            moveDefinition.Add(new ModuleDefinition("Module3", ComponentType.StandardModule, externalReferencingCode));


            var refactoredCode = ExecuteTest(moveDefinition);

            StringAssert.AreEqualIgnoringCase(ThisStrategy, refactoredCode.StrategyName);

            StringAssert.DoesNotContain("Sub Foo(arg1", refactoredCode.Source);

            StringAssert.Contains($"{accessibility} Sub Foo(arg1", refactoredCode.Destination);
            StringAssert.DoesNotContain("Public Property Let Bar", refactoredCode.Destination);
            StringAssert.DoesNotContain("Public Property Get Bar", refactoredCode.Destination);

            var module3Content = refactoredCode["Module3"];
            StringAssert.Contains($"{moveDefinition.SourceModuleName}.Bar", refactoredCode.Destination);

        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void PlacesCodeInCorrectSpotRelativeToDeclareStmt()
        {
            var memberToMove = ("Fizz", ThisDeclarationType);
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

            var moveDefinition = new TestMoveDefinition(MoveEndpoints.StdToStd, memberToMove);
            moveDefinition.SetEndpointContent(source, destination);

            var refactoredCode = ExecuteTest(moveDefinition);

            StringAssert.AreEqualIgnoringCase(ThisStrategy, refactoredCode.StrategyName);

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
            var memberToMove = ("RaisesAnEvent", ThisDeclarationType);
            var source = $@"
Option Explicit

Public Event EventName(IDNumber As Long, ByRef Cancel As Boolean)

Public Sub RaisesAnEvent()
    RaiseEvent EventName(6, true)
End Sub";

            var eventSinkName = "CEventSink";
            var eventSinkContent = $@"
Option Explicit

Dim WithEvents TestEvents As {Support.DEFAULT_SOURCE_CLASS_NAME}

Private Sub TestEvents_EventName(IDNumber As Long, ByRef Cancel As Boolean)
End Sub
";

            var moveDefinition = new TestMoveDefinition(MoveEndpoints.ClassToStd, memberToMove, source);
            moveDefinition.Add(new ModuleDefinition(eventSinkName, ComponentType.ClassModule, eventSinkContent));

            var refactoredCode = ExecuteTest(moveDefinition);

            Assert.IsNull(refactoredCode.StrategyName);
        }


        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void DoesNotMoveEventSink()
        {
            var eventSourceName = "CEventSource";

            var memberToMove = ("TestEvents_EventName", ThisDeclarationType);
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

            var moveDefinition = new TestMoveDefinition(MoveEndpoints.ClassToStd, memberToMove, source);
            moveDefinition.Add(new ModuleDefinition(eventSourceName, ComponentType.ClassModule, eventSourceContent));

            var refactoredCode = ExecuteTest(moveDefinition);

            Assert.IsNull(refactoredCode.StrategyName);
        }

        [Test]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void DoesNotMoveInterfaceImplementationMember()
        {
            var memberToMove = ("ITestInterface_TestGet", DeclarationType.PropertyGet);
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

            var moveDefinition = new TestMoveDefinition(MoveEndpoints.ClassToStd, memberToMove, source);
            moveDefinition.Add(new ModuleDefinition(interfaceDeclarationClass, ComponentType.ClassModule, interfaceContent));

            var refactoredCode = ExecuteTest(moveDefinition);

            Assert.IsNull(refactoredCode.StrategyName);
        }

    }
}
