using Rubberduck.Parsing.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.ComCommands
{
    public class ProjectExplorerIgnoreProjectCommand : ComCommandBase
    {
        private readonly IVBE _vbe;
        private readonly IConfigurationService<IgnoredProjectsSettings> _configService;

        public ProjectExplorerIgnoreProjectCommand(IVbeEvents vbeEvents, IVBE vbe, IConfigurationService<IgnoredProjectsSettings> configService) 
            : base(vbeEvents)
        {
            _vbe = vbe;
            _configService = configService;
            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute, true);
        }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            using (var activeProject = _vbe.ActiveVBProject)
            {
                if (activeProject == null 
                    || !activeProject.TryGetFullPath(out var fullPath))
                {
                    return false;
                }

                var ignoredProjectPaths = _configService.Read().IgnoredProjectPaths;
                return !ignoredProjectPaths.Contains(fullPath);
            }
        }

        protected override void OnExecute(object parameter)
        {
            using (var activeProject = _vbe.ActiveVBProject)
            {
                if (activeProject == null
                    || !activeProject.TryGetFullPath(out var fullPath))
                {
                    return;
                }

                var ignoredProjectsSetting = _configService.Read();
                if (!ignoredProjectsSetting.IgnoredProjectPaths.Contains(fullPath))
                {
                    ignoredProjectsSetting.IgnoredProjectPaths.Add(fullPath);
                    _configService.Save(ignoredProjectsSetting);
                }
            }
        }
    }
}