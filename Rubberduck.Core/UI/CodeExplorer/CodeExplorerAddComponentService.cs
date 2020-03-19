using System;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Text;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.CodeExplorer
{
    public class CodeExplorerAddComponentService : ICodeExplorerAddComponentService
    {            
        private readonly IParseManager _parseManager;
        private readonly IAddComponentService _addComponentService;
        private readonly IVBE _vbe;

        public CodeExplorerAddComponentService(IParseManager parseManager, IAddComponentService addComponentService, IVBE vbe)
        {
            _parseManager = parseManager;
            _addComponentService = addComponentService;
            _vbe = vbe;
        }

        public void AddComponent(CodeExplorerItemViewModel node, ComponentType componentType, string code = null)
        {
            var projectId = ProjectId(node);
            if (projectId == null)
            {
                return;
            }

            var prefixInModule = FolderAnnotation(node);

            var suspensionResult = _parseManager.OnSuspendParser(
                this,
                Enum.GetValues(typeof(ParserState)).Cast<ParserState>(),
                () => _addComponentService.AddComponent(projectId, componentType, code, prefixInModule));

            if (suspensionResult.Outcome == SuspensionOutcome.UnexpectedError 
                && suspensionResult.EncounteredException != null)
            {
                //This rethrows with the original stack trace.
                ExceptionDispatchInfo.Capture(suspensionResult.EncounteredException).Throw();
            }
        }

        public void AddComponentWithAttributes(CodeExplorerItemViewModel node, ComponentType componentType, string code, string additionalPrefixInModule = null)
        {
            var projectId = ProjectId(node);
            if (projectId == null)
            {
                return;
            }

            var folderAnnotation = FolderAnnotation(node);
            var optionCompare = OptionCompareStatement();

            var modulePrefix = new StringBuilder(folderAnnotation);
            if (optionCompare != null)
            {
                modulePrefix.Append(Environment.NewLine).Append(optionCompare);
            }
            if (additionalPrefixInModule != null)
            {
                modulePrefix.Append(Environment.NewLine).Append(additionalPrefixInModule);
            }
            var prefixInModule = modulePrefix.ToString();

            var suspensionResult = _parseManager.OnSuspendParser(
                this,
                Enum.GetValues(typeof(ParserState)).Cast<ParserState>(),
                () => _addComponentService.AddComponentWithAttributes(projectId, componentType, code, prefixInModule));

            if (suspensionResult.Outcome == SuspensionOutcome.UnexpectedError
                && suspensionResult.EncounteredException != null)
            {
                //This rethrows with the original stack trace.
                ExceptionDispatchInfo.Capture(suspensionResult.EncounteredException).Throw();
            }
        }

        private string ProjectId(CodeExplorerItemViewModel node)
        {
            return node?.Declaration.ProjectId;
        }

        private string FolderAnnotation(CodeExplorerItemViewModel node)
        {
            return (node is CodeExplorerCustomFolderViewModel folder) 
                ? folder.FolderAttribute 
                : $"'@Folder(\"{Folder(node)}\")";
        }

        private string Folder(CodeExplorerItemViewModel node)
        {
            var declaration = node?.Declaration;
            if (declaration == null)
            {
                return ActiveProjectFolder();
            }

            return declaration.CustomFolder ?? ProjectFolder(declaration.ProjectName);
        }

        private string ActiveProjectFolder()
        {
            return ProjectFolder(ActiveProjectName());
        }

        private string ActiveProjectName()
        {
            using (var activeProject = _vbe.ActiveVBProject)
            {
                return activeProject?.Name;
            }
        }

        private static string ProjectFolder(string projectName)
        {
            return projectName;
        }

        private string OptionCompareStatement()
        {
            using (var hostApp = _vbe.HostApplication())
            {
                return hostApp?.ApplicationName == "Access" 
                    ? "Option Compare Database" 
                    : null;
            }
        }
    }
}