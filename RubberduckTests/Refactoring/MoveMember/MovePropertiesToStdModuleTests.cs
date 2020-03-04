using NUnit.Framework;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.MoveMember;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Support = RubberduckTests.Refactoring.MoveMember.MoveMemberTestSupport;

namespace RubberduckTests.Refactoring.MoveMember
{
    [TestFixture]
    public class MovePropertiesToStdModuleTests : MoveMemberTestsBase
    {
        private const string ThisStrategy = nameof(MoveMemberToStdModule);

        [TestCase(MoveEndpoints.StdToStd)]
        [TestCase(MoveEndpoints.ClassToStd)]
        [TestCase(MoveEndpoints.FormToStd)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void PreviewMovedContent(MoveEndpoints endpoints)
        {
            var memberToMove = ("TheValue", DeclarationType.PropertyGet);
            var source =
$@"
Option Explicit


Private mTheValue As Long

Public Property Get TheValue() As Long
    TheValue = mTheValue
End Property

Public Property Let TheValue(ByVal value As Long)
    mTheValue = value
End Property
";

            var moveDefinition = new TestMoveDefinition(endpoints, memberToMove);
            var preview = RetrievePreviewAfterUserInput(moveDefinition, source, memberToMove);

            StringAssert.Contains("Option Explicit", preview);
            Assert.IsTrue(Support.OccursOnce("Property Get TheValue(", preview));
            Assert.IsTrue(Support.OccursOnce("Property Let TheValue(", preview));
        }


        [TestCase(MoveEndpoints.StdToStd, "Public", "Private Const Pi As Single = 3.14", ThisStrategy)]
        [TestCase(MoveEndpoints.StdToStd, "Public", "Private Pi As Single", ThisStrategy)]
        [TestCase(MoveEndpoints.StdToStd, "Private", "Private Const Pi As Single = 3.14", ThisStrategy)]
        [TestCase(MoveEndpoints.StdToStd, "Private", "Private Pi As Single", ThisStrategy)]
        [TestCase(MoveEndpoints.ClassToStd, "Public", "Private Const Pi As Single = 3.14", ThisStrategy)]
        [TestCase(MoveEndpoints.ClassToStd, "Public", "Private Pi As Single", ThisStrategy)]
        [TestCase(MoveEndpoints.ClassToStd, "Private", "Private Const Pi As Single = 3.14", ThisStrategy)]
        [TestCase(MoveEndpoints.ClassToStd, "Private", "Private Pi As Single", ThisStrategy)]
        [TestCase(MoveEndpoints.FormToStd, "Public", "Private Const Pi As Single = 3.14", ThisStrategy)]
        [TestCase(MoveEndpoints.FormToStd, "Public", "Private Pi As Single", ThisStrategy)]
        [TestCase(MoveEndpoints.FormToStd, "Private", "Private Const Pi As Single = 3.14", ThisStrategy)]
        [TestCase(MoveEndpoints.FormToStd, "Private", "Private Pi As Single", ThisStrategy)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void SupportFunctionsReferenceExclusiveSupportNonMember(MoveEndpoints endpoints, string exclusiveFuncAccessibility, string nonMemberDeclaration, string expectedStrategy)
        {
            //Using a global variable to isolate the test to the supporting constant 'Pi'
            var moduleRadius = "RadiusModule";
            var moduleRadiusContent =
@"Option Explicit

Public Radius As Single
";
            var memberToMove = "Area";
            var source =
$@"
Option Explicit

{nonMemberDeclaration}

Public Property Get Area() As Single
    Area = ToArea({moduleRadius}.Radius)
End Property

Public Property Let Area(area As Single)
    {moduleRadius}.Radius = ToRadius(area) 
End Property

{exclusiveFuncAccessibility} Function ToRadius(area As Single) As Single
    ToRadius = Sqr(area / Pi)
End Function

{exclusiveFuncAccessibility} Function ToArea(radius As Single) As Single
    ToArea = Pi * {moduleRadius}.Radius ^ 2
End Function
";

            var destinationExpectedForClassAndFormModules =
$@"
Option Explicit

{nonMemberDeclaration}

Public Property Get Area() As Single
    Area = ToArea({moduleRadius}.Radius)
End Property

Public Property Let Area(ByVal area As Single)
    {moduleRadius}.Radius = ToRadius(area) 
End Property

{exclusiveFuncAccessibility} Function ToRadius(ByRef area As Single) As Single
    ToRadius = Sqr(area / Pi)
End Function

{exclusiveFuncAccessibility} Function ToArea(ByRef radius As Single) As Single
    ToArea = Pi * {moduleRadius}.Radius ^ 2
End Function
";

            var moveDefinition = new TestMoveDefinition(endpoints, (memberToMove, DeclarationType.PropertyGet));
            moveDefinition.Add(new ModuleDefinition(moduleRadius, Rubberduck.VBEditor.SafeComWrappers.ComponentType.StandardModule, moduleRadiusContent));

            var refactoredCode = RefactoredCode_UserSetsDestinationModuleName(moveDefinition, source);

            StringAssert.AreEqualIgnoringCase(expectedStrategy, refactoredCode.StrategyName);

            if (expectedStrategy is null)
            {
                return;
            }

            if (!moveDefinition.IsStdModuleSource)
            {
                var refactoredLines = refactoredCode.Destination.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var line in refactoredLines)
                {
                    //Moves everything from Source to Destination as-is
                    StringAssert.Contains(line, destinationExpectedForClassAndFormModules);
                }
                return;
            }

            //Source Module content checks
            StringAssert.DoesNotContain("Get Area() As Single", refactoredCode.Source);
            StringAssert.DoesNotContain("Let Area(area As Single)", refactoredCode.Source);
            if (exclusiveFuncAccessibility.Equals(Tokens.Private))
            {
                StringAssert.DoesNotContain($"{exclusiveFuncAccessibility} Function ToArea(", refactoredCode.Source);
                StringAssert.DoesNotContain($"{exclusiveFuncAccessibility} Function ToRadius(", refactoredCode.Source);
                StringAssert.DoesNotContain(nonMemberDeclaration, refactoredCode.Source);
            }
            else
            {
                StringAssert.Contains($"{exclusiveFuncAccessibility} Function ToArea(", refactoredCode.Source);
                StringAssert.Contains($"{exclusiveFuncAccessibility} Function ToRadius(", refactoredCode.Source);
                StringAssert.Contains(nonMemberDeclaration, refactoredCode.Source);
            }

            //Destination module content checks
            if (exclusiveFuncAccessibility.Equals(Tokens.Private))
            {
                StringAssert.Contains(nonMemberDeclaration, refactoredCode.Destination);
                StringAssert.Contains($"{exclusiveFuncAccessibility} Function ToArea(", refactoredCode.Destination);
                StringAssert.Contains($"{exclusiveFuncAccessibility} Function ToRadius(", refactoredCode.Destination);
                StringAssert.Contains($"{moduleRadius}.Radius = ToRadius(area)", refactoredCode.Destination);
            }
            else
            {
                StringAssert.DoesNotContain(nonMemberDeclaration, refactoredCode.Destination);
                StringAssert.DoesNotContain($"{exclusiveFuncAccessibility} Function ToArea(", refactoredCode.Destination);
                StringAssert.DoesNotContain($"{exclusiveFuncAccessibility} Function ToRadius(", refactoredCode.Destination);
                StringAssert.Contains($"Area = {moveDefinition.SourceModuleName}.ToArea(", refactoredCode.Destination);
                StringAssert.Contains($"{moduleRadius}.Radius = {moveDefinition.SourceModuleName}.ToRadius(", refactoredCode.Destination);
            }
            StringAssert.Contains("Public Property Get Area(", refactoredCode.Destination);
            StringAssert.Contains("Public Property Let Area(", refactoredCode.Destination);
        }
    }
}
