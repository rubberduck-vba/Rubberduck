using System.Linq;
using System.Windows.Forms;
using Rubberduck.AddRemoveReferences;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.AddRemoveReferences
{
    public interface IAddRemoveReferencesPresenter
    {
        IAddRemoveReferencesModel Show();
        IAddRemoveReferencesModel Show(ProjectDeclaration project);
        IAddRemoveReferencesModel Model { get; }
    }

    public class AddRemoveReferencesPresenter : IAddRemoveReferencesPresenter
    {
        private readonly AddRemoveReferencesDialog _view;

        public AddRemoveReferencesPresenter(AddRemoveReferencesDialog view)
        {
            _view = view;
            Model = _view.ViewModel.Model;
        }

        public IAddRemoveReferencesModel Show()
        {
            return Model.Project == null ? null : Show(Model.Project);
        }

        public IAddRemoveReferencesModel Show(ProjectDeclaration project)
        {
            if (project is null)
            {
                return null;
            }

            Model.Project = project;
            _view.ViewModel.Model = Model;

            _view.ShowDialog();
            if (_view.DialogResult != DialogResult.OK)
            {
                return null;
            }

            Model.NewReferences = _view.ViewModel.ProjectReferences.SourceCollection.OfType<ReferenceModel>().ToList(); 
            return Model;
        }

        public IAddRemoveReferencesModel Model { get; }
    }
}
