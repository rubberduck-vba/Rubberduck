using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class UpdateFromFilesCommand : ImportCommand
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IProjectsProvider _projectsProvider;
        private readonly IModuleNameFromFileExtractor _moduleNameFromFileExtractor;

        public UpdateFromFilesCommand(
            IVBE vbe,
            IFileSystemBrowserFactory dialogFactory,
            IVbeEvents vbeEvents,
            IParseManager parseManager,
            IDeclarationFinderProvider declarationFinderProvider,
            IProjectsProvider projectsProvider,
            IModuleNameFromFileExtractor moduleNameFromFileExtractor)
            : base(vbe, dialogFactory, vbeEvents, parseManager)
        {
            _projectsProvider = projectsProvider;
            _declarationFinderProvider = declarationFinderProvider;
            _moduleNameFromFileExtractor = moduleNameFromFileExtractor;
        }

        protected override void ImportFiles(ICollection<string> filesToImport, IVBProject targetProject)
        {
            var finder = _declarationFinderProvider.DeclarationFinder;

            var moduleNames = ModuleNames(filesToImport);

            if (!ValuesAreUnique(moduleNames))
            {
                //TODO: report this to the user.
                return;
            }

            var modules = Modules(moduleNames, targetProject.ProjectId, finder);

            //TODO: abort if the component type of the to be removed component does not match the file extension.

            using (var components = targetProject.VBComponents)
            {
                foreach (var filename in filesToImport)
                {
                    if (modules.TryGetValue(filename, out var module))
                    {
                        var component = _projectsProvider.Component(module);
                        components.Remove(component);
                    }

                    //We have to dispose the return value.
                    using (components.Import(filename)) { }
                }
            }
        }

        private Dictionary<string, string> ModuleNames(ICollection<string> filenames)
        {
            var moduleNames = new Dictionary<string, string>();
            foreach(var filename in filenames)
            {
                if (moduleNames.ContainsKey(filename))
                {
                    continue;
                }

                var moduleName = ModuleName(filename);
                if(moduleName != null)
                {
                    moduleNames.Add(filename, moduleName);
                }
            }

            return moduleNames;
        }

        private string ModuleName(string filename)
        {
            return _moduleNameFromFileExtractor.ModuleName(filename);
        }

        private Dictionary<string, QualifiedModuleName> Modules(IDictionary<string, string> moduleNames, string projectId, DeclarationFinder finder)
        {
            var modules = new Dictionary<string, QualifiedModuleName>();
            foreach (var (fileName, moduleName) in moduleNames)
            {
                var module = Module(moduleName, projectId, finder);
                if (module.HasValue)
                {
                    modules.Add(fileName, module.Value);
                }
            }

            return modules;
        }

        private bool ValuesAreUnique(Dictionary<string, string> moduleNames)
        {
            return moduleNames
                .GroupBy(kvp => kvp.Value)
                .All(moduleNameGroup => moduleNameGroup.Count() == 1);
        }

        private QualifiedModuleName? Module(string moduleName, string projectId, DeclarationFinder finder)
        {
            foreach(var module in finder.AllModules)
            {
                if(module.ProjectId.Equals(projectId)
                    && module.ComponentName.Equals(moduleName))
                {
                    return module;
                }
            }

            return null;
        }
    }

    public interface IModuleNameFromFileExtractor
    {
        string ModuleName(string filename);
    }

    public class ModuleNameFromFileExtractor : IModuleNameFromFileExtractor
    {
        public string ModuleName(string filename)
        {
            if (!File.Exists(filename))
            {
                return null;
            }

            var contents = File.ReadLines(filename, Encoding.Default);
            var nameLine = contents.FirstOrDefault(line => line.StartsWith("Attribute VB_Name = "));
            if (nameLine == null)
            {
                return Path.GetFileNameWithoutExtension(filename);
            }

            //The format is Attribute VB_Name = "ModuleName"
            return nameLine.Substring("Attribute VB_Name = ".Length + 1, nameLine.Length - "Attribute VB_Name = ".Length - 2);
        }
    }
}