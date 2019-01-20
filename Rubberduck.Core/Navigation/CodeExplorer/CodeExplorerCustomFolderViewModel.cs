using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Rubberduck.Navigation.Folders;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Navigation.CodeExplorer
{
    [DebuggerDisplay("{Name}")]
    public sealed class CodeExplorerCustomFolderViewModel : CodeExplorerItemViewModel
    {
        private static readonly DeclarationType[] ComponentTypes =
        {
            DeclarationType.ClassModule, 
            DeclarationType.Document, 
            DeclarationType.ProceduralModule, 
            DeclarationType.UserForm
        };

        private readonly IVBE _vbe;

        public CodeExplorerCustomFolderViewModel(
            ICodeExplorerNode parent, 
            string name, 
            string fullPath, 
            IVBE vbe,
            IEnumerable<Declaration> declarations) : base(parent, parent?.Declaration)
        {
            _vbe = vbe;
            FolderDepth = parent is CodeExplorerCustomFolderViewModel folder ? folder.FolderDepth + 1 : 1;
            FullPath = fullPath?.Trim('"') ?? string.Empty;
            Name = name.Replace("\"", string.Empty);

            AddNewChildren(declarations.ToList());
        }

        public override string Name { get; }

        public override string PanelTitle => FullPath ?? string.Empty;

        public override string Description => FolderAttribute ?? string.Empty;

        public string FullPath { get; }

        public string FolderAttribute => $"'@Folder(\"{FullPath.Replace("\"", string.Empty)}\")";

        /// <summary>
        /// One-based depth in the folder hierarchy.
        /// </summary>
        public int FolderDepth { get; }

        public override QualifiedSelection? QualifiedSelection => null;

        public override bool IsErrorState
        {
            get => false;
            set { /* Folders can never be in an error state. */ }
        }

        public override Comparer<ICodeExplorerNode> SortComparer => CodeExplorerItemComparer.Name;

        public override bool Filtered => false;

        protected override void AddNewChildren(List<Declaration> declarations)
        {
            var children = declarations.Where(declaration => declaration.IsInSubFolder(FullPath)).ToList();

            foreach (var folder in children.GroupBy(declaration => declaration.CustomFolder.SubFolderRoot(FullPath)))
            {
                AddChild(new CodeExplorerCustomFolderViewModel(this, folder.Key, $"{FullPath}.{folder.Key}", _vbe, folder));
                foreach (var declaration in folder)
                {
                    declarations.Remove(declaration);
                }
            }

            foreach (var declaration in declarations.Where(child => child.IsInFolder(FullPath)).GroupBy(item => item.ComponentName))
            {
                var moduleName = declaration.Key;
                var parent = declarations.SingleOrDefault(item => 
                    ComponentTypes.Contains(item.DeclarationType) && item.ComponentName == moduleName);

                if (parent is null)
                {
                    continue;
                }

                var members = declarations.Where(item =>
                    !ComponentTypes.Contains(item.DeclarationType) && item.ComponentName == moduleName);

                AddChild(new CodeExplorerComponentViewModel(this, parent, members, _vbe));
                declarations.Remove(parent);
            }
        }

        public override void Synchronize(List<Declaration> updated)
        {
            SynchronizeChildren(updated);
        }

        protected override void SynchronizeChildren(List<Declaration> updated)
        {
            var declarations = updated.Where(declaration => declaration.IsInFolderOrSubFolder(FullPath)).ToList();

            if (!declarations.Any())
            {
                Declaration = null;
                return;
            }

            var synchronizing = declarations.ToList();

            foreach (var subfolder in Children.OfType<CodeExplorerCustomFolderViewModel>().ToList())
            {
                subfolder.SynchronizeChildren(declarations);
                if (subfolder.Declaration is null)
                {
                    RemoveChild(subfolder);
                    continue;
                }

                var synchronized = synchronizing.Where(child => !declarations.Contains(child)).ToList();
                foreach (var declaration in synchronized)
                {
                    updated.Remove(declaration);
                }
            }

            foreach (var child in Children.OfType<CodeExplorerComponentViewModel>().ToList())
            {
                child.Synchronize(updated);
                if (child.Declaration is null)
                {
                    RemoveChild(child);
                    continue;
                }

                updated.Remove(child.Declaration);
            }

            AddNewChildren(updated);
        }
    }
}
