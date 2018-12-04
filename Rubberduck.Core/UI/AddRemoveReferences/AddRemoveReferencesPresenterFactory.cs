using System;
using System.Collections.Generic;
using System.Linq;
using NLog;
using Rubberduck.AddRemoveReferences;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.AddRemoveReferences
{
    public interface IAddRemoveReferencesPresenterFactory : IRefactoringPresenterFactory<AddRemoveReferencesPresenter>
    {
        AddRemoveReferencesPresenter Create(ProjectDeclaration project);
    }

    public class AddRemoveReferencesPresenterFactory : IAddRemoveReferencesPresenterFactory
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly bool _use64BitPaths = Environment.Is64BitProcess;

        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IConfigProvider<GeneralSettings> _settings;
        private readonly IRegisteredLibraryFinderService _finder;
        private readonly IMessageBox _messageBox;

        public AddRemoveReferencesPresenterFactory(IVBE vbe,
            RubberduckParserState state,
            IConfigProvider<GeneralSettings> generalSettingsProvider, 
            IRegisteredLibraryFinderService finder,
            IMessageBox messageBox)
        {
            _vbe = vbe;
            _state = state;
            _settings = generalSettingsProvider;
            _finder = finder;
            _messageBox = messageBox;
        }

        public AddRemoveReferencesPresenter Create(ProjectDeclaration project)
        {
            if (project is null)
            {
                return null;
            }

            var refs = new Dictionary<RegisteredLibraryKey, RegisteredLibraryInfo>();
            // Iterating the returned libraries here instead of just .ToDictionary() using because we can't trust that the registry doesn't contain errors.
            foreach (var reference in _finder.FindRegisteredLibraries())
            {
                if (refs.ContainsKey(reference.UniqueId))
                {
                    _logger.Warn($"Duplicate registry definition for {reference.Guid} version {reference.Version}.");
                    continue;
                }
                refs.Add(reference.UniqueId, reference);
            }

            var models = new Dictionary<RegisteredLibraryKey, ReferenceModel>();
            using (var references = project.Project?.References)
            {
                if (references is null)
                {
                    return null;
                }
                var priority = 1;
                foreach (var reference in references)
                {
                    var libraryId = new RegisteredLibraryKey(new Guid(reference.Guid), reference.Major, reference.Minor);
                    if (refs.ContainsKey(libraryId))
                    {
                        // TODO: If for some reason the VBA reference is broken, we could technically use this to repair it. Just a thought...
                        models.Add(libraryId, new ReferenceModel(refs[libraryId], reference, priority++));
                    }
                    else // These should all be either VBA projects or irreparably broken.
                    {
                        models.Add(libraryId, new ReferenceModel(reference, priority++));
                    }
                    reference.Dispose();
                }
            }

            foreach (var reference in refs.Where(library =>
                (_use64BitPaths || library.Value.Has32BitVersion) &&
                !models.ContainsKey(library.Key)))
            {
                models.Add(reference.Key, new ReferenceModel(reference.Value));
            }

            var settings = _settings.Create();
            var model = new AddRemoveReferencesModel(project, models.Values, settings);

            return new AddRemoveReferencesPresenter(new AddRemoveReferencesDialog(new AddRemoveReferencesViewModel(model, _messageBox)));         
        }

        public AddRemoveReferencesPresenter Create()
        {
            using (var pane = _vbe.ActiveCodePane)
            {
                var selected = (ProjectDeclaration)Declaration.GetProjectParent(_state.DeclarationFinder.FindSelectedDeclaration(pane));
                return selected is null ? null : Create(selected);
            }
        }
    }
}
