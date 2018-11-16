﻿using System.Linq;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Navigation.Folders
{
    public class FolderHelper
    {
        private readonly RubberduckParserState _state;
        private readonly IVBE _vbe;

        private static readonly DeclarationType[] ComponentTypes =
        {
            DeclarationType.ClassModule, 
            DeclarationType.Document, 
            DeclarationType.ProceduralModule, 
            DeclarationType.UserForm, 
        };

        public FolderHelper(RubberduckParserState state, IVBE vbe)
        {
            _state = state;
            _vbe = vbe;
        }

        public CodeExplorerCustomFolderViewModel GetFolderTree(Declaration declaration = null)
        {
            var delimiter = GetDelimiter();

            var root = new CodeExplorerCustomFolderViewModel(null, string.Empty, string.Empty, _state.ProjectsProvider, _vbe);

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
                        node = new CodeExplorerCustomFolderViewModel(currentNode, section, fullPath, _state.ProjectsProvider, _vbe);
                        currentNode.AddChild(node);
                    }

                    currentNode = (CodeExplorerCustomFolderViewModel)node;
                }
            }

            return root;
        }

        private char GetDelimiter() => '.';
    }
}
