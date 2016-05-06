using System.Collections.Generic;
using System.Linq;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;

namespace Rubberduck.Navigation.Folders
{
    public class FolderHelper
    {
        private readonly RubberduckParserState _state;
        private readonly ConfigurationLoader _configLoader;

        private static readonly DeclarationType[] ComponentTypes =
        {
            DeclarationType.ClassModule, 
            DeclarationType.Document, 
            DeclarationType.ProceduralModule, 
            DeclarationType.UserForm, 
        };

        public FolderHelper(RubberduckParserState state, ConfigurationLoader configLoader)
        {
            _state = state;
            _configLoader = configLoader;
        }

        public CodeExplorerCustomFolderViewModel GetFolderTree()
        {
            var delimiter = GetDelimiter();

            var root = new CodeExplorerCustomFolderViewModel(string.Empty, string.Empty);

            var items = _state.AllUserDeclarations.ToList();
            var folders = items.Where(item => ComponentTypes.Contains(item.DeclarationType))
                .Select(item => item.CustomFolder.Replace("\"", string.Empty))
                .Distinct()
                .Select(item => item.Split(delimiter));

            foreach (var folder in folders)
            {
                var currentNode = root;
                var fullPath = string.Empty;

                foreach (var section in folder)
                {
                    fullPath += string.IsNullOrEmpty(fullPath) ? section : delimiter + section;

                    var node = currentNode.Items.FirstOrDefault(i => i.Name == section);
                    if (node == null)
                    {
                        node = new CodeExplorerCustomFolderViewModel(section, fullPath);
                        currentNode.AddChild(node);
                    }

                    currentNode = (CodeExplorerCustomFolderViewModel)node;
                }
            }

            return root;
        }

        private char GetDelimiter()
        {
            var settings = _configLoader.LoadConfiguration();
            return settings.UserSettings.GeneralSettings.Delimiter;
        }
    }
}