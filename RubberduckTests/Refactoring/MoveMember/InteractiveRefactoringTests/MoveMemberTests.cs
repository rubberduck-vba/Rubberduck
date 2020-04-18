using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RubberduckTests.Refactoring.MoveMember
{
    [TestFixture]
    public class MoveMemberTests : InteractiveRefactoringTestBase<IMoveMemberPresenter, MoveMemberModel>
    {
        [TestCase(MoveEndpoints.ClassToStd)]
        [TestCase(MoveEndpoints.FormToStd)]
        [TestCase(MoveEndpoints.StdToStd)]
        [Category("Refactorings")]
        [Category("MoveMember")]
        public void MoveToEmptyDestinationSelectAll(MoveEndpoints endpoints)
        {
            var sourceModule = endpoints.SourceModuleName();
            var destinationModule = endpoints.DestinationModuleName();
            var source =
$@"
Option Explicit

Private mTizz As Long

Function Fi|zz(arg1 As Long) As Long
End Function

Sub Hizz(arg1 As Long)
End Sub

Property Get Tizz() As Long
    Tizz = mTizz
End Property

Property Let Tizz(value As Long)
    mTizz = value
End Property
";
            Func<MoveMemberModel, MoveMemberModel> presenterAction = model =>
            {
                foreach (var moveableMemberSet in model.MoveableMemberSets)
                {
                    moveableMemberSet.IsSelected = true;
                }
                model.ChangeDestination(endpoints.DestinationModuleName(), endpoints.DestinationComponentType());
                return model;
            };

            var input  =  ToSelectionAndCode(source.ToCodeString());
            var actualCode = RefactoredCode(sourceModule, input.Selection, presenterAction, null, false, endpoints.ToSourceTuple(input.Code), endpoints.ToDestinationTuple(string.Empty));

            StringAssert.Contains("Private mTizz", actualCode[destinationModule]);
            StringAssert.Contains("Public Function Fizz(", actualCode[destinationModule]);
            StringAssert.Contains("Public Sub Hizz(", actualCode[destinationModule]);
            StringAssert.Contains("Public Property Get Tizz", actualCode[destinationModule]);
            StringAssert.Contains("Public Property Let Tizz", actualCode[destinationModule]);
        }

        protected override IRefactoring TestRefactoring(IRewritingManager rewritingManager, RubberduckParserState state, RefactoringUserInteraction<IMoveMemberPresenter, MoveMemberModel> userInteraction, ISelectionService selectionService)
        {
            return MoveMemberTestsResolver.CreateRefactoring(rewritingManager, state, userInteraction, selectionService);
        }

        private (Selection Selection, string Code) ToSelectionAndCode(CodeString input)
        {
            return (input.CaretPosition.ToOneBased(), input.Code);
        }
    }
}
