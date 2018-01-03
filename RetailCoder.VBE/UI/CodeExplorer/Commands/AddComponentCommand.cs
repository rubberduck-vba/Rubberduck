﻿using System.Runtime.InteropServices;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class AddComponentCommand
    {
        private readonly IVBE _vbe;
        private const string DefaultFolder = "VBAProject";

        public AddComponentCommand(IVBE vbe)
        {
            _vbe = vbe;
        }

        public bool CanAddComponent(CodeExplorerItemViewModel parameter)
        {
            try
            {
                return GetDeclaration(parameter) != null || _vbe.ProjectsCount == 1;
            }
            catch (COMException)
            {
                return false;
            }
        }

        public void AddComponent(CodeExplorerItemViewModel node, ComponentType type)
        {
            var components = node != null
                ? GetDeclaration(node).Project.VBComponents
                : _vbe.ActiveVBProject.VBComponents;

            var folderAnnotation = $"'@Folder(\"{GetFolder(node)}\")";

            var newComponent = components.Add(type);
            newComponent.CodeModule.InsertLines(1, folderAnnotation);
        }

        private Declaration GetDeclaration(CodeExplorerItemViewModel node)
        {
            while (node != null && !(node is ICodeExplorerDeclarationViewModel))
            {
                node = node.Parent;
            }

            return (node as ICodeExplorerDeclarationViewModel)?.Declaration;
        }

        private string GetFolder(CodeExplorerItemViewModel node)
        {
            switch (node)
            {
                case null:
                    return DefaultFolder;
                case ICodeExplorerDeclarationViewModel declarationNode:
                    return string.IsNullOrEmpty(declarationNode.Declaration.CustomFolder)
                        ? DefaultFolder
                        : declarationNode.Declaration.CustomFolder.Replace("\"", string.Empty);
                default:
                    return ((CodeExplorerCustomFolderViewModel)node).FullPath;
            }
        }
    }
}