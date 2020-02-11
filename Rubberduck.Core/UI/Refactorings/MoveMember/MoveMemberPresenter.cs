using Rubberduck.Refactorings;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.VBEditor.SafeComWrappers;
using System.Collections.Generic;

namespace Rubberduck.UI.Refactorings.MoveMember
{
    public class MoveMemberPresenter : RefactoringPresenterBase<MoveMemberModel>, IMoveMemberPresenter
    {
        private readonly MoveMemberViewModel _view;

        private static readonly DialogData DialogData = DialogData.Create(MoveMemberResources.Caption, minimumHeight: 450, minimumWidth: 800);

        public MoveMemberPresenter(MoveMemberModel model, IRefactoringDialogFactory dialogFactory)
            : base(DialogData, model, dialogFactory)
        {
            _view = (MoveMemberViewModel)ViewModel;
        }

        public override MoveMemberModel Show()
        {
            if (Model != null)
            {
                _view.SourceModule = Model.Source.Module;
                _view.MemberToMove = Model.DefiningMember;
            }

            _view.Preview = GetNewModuleContent;
            _view.ConflictsRetriever = GetMoveConflictDescriptors;

            var model = base.Show();

            if (DialogResult == RefactoringDialogResult.Execute)
            {
                if (_view.DestinationModule != null)
                {
                    model?.DefineMove(_view.MemberToMove, _view.DestinationModule);
                }
                else
                {
                    model?.DefineMove(_view.MemberToMove, destinationModuleName: _view.DestinationModuleName, destinationType: ComponentType.StandardModule);
                }
                return model;
            }
            return null;
        }

        private string GetNewModuleContent()
        {
            if (Model != null)
            {
                DefineMove();
                return Model.PreviewDestination();
            }
            return string.Empty;
        }

        private IEnumerable<string> GetMoveConflictDescriptors()
        {
            DefineMove();
            return Model?.HasValidDestination ?? false ? new List<string>() : new List<string>() { _view.Instructions };
        }

        private void DefineMove()
        {
            if (_view.DestinationModule != null)
            {
                Model?.DefineMove(_view.MemberToMove, _view.DestinationModule);
            }
            else
            {
                Model?.DefineMove(_view.MemberToMove, _view.DestinationModuleName, ComponentType.StandardModule);
            }
        }
    }
}
