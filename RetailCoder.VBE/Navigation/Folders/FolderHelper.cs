using System.Linq;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Navigation.Folders
{
    public class FolderHelper
    {
        private readonly RubberduckParserState _state;

        private static readonly DeclarationType[] ComponentTypes =
        {
            DeclarationType.ClassModule, 
            DeclarationType.Document, 
            DeclarationType.ProceduralModule, 
            DeclarationType.UserForm, 
        };

        public FolderHelper(RubberduckParserState state)
        {
            _state = state;
        }

        public CodeExplorerCustomFolderViewModel GetFolderTree(Declaration declaration = null)
        {
            var delimiter = GetDelimiter();

            var root = new CodeExplorerCustomFolderViewModel(null, string.Empty, string.Empty);

            var items = declaration == null
                ? _state.AllUserDeclarations.ToList()
                : _state.AllUserDeclarations.Where(d => d.ProjectId == declaration.ProjectId).ToList();

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
                        node = new CodeExplorerCustomFolderViewModel(currentNode, section, fullPath);
                        currentNode.AddChild(node);
                    }

                    currentNode = (CodeExplorerCustomFolderViewModel)node;
                }
            }

            return root;
        }

        private char GetDelimiter()
        {
            return '.';
        }
    }
}
