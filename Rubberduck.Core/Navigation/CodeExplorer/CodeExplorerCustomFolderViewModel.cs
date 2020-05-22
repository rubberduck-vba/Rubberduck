using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.JunkDrawer.Extensions;
using Rubberduck.Navigation.Folders;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Navigation.CodeExplorer
{
    [DebuggerDisplay("{" + nameof(Name) + "}")]
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
            ref List<Declaration> declarations) : base(parent, parent?.Declaration)
        {
            _vbe = vbe;
            FolderDepth = parent is CodeExplorerCustomFolderViewModel folder ? folder.FolderDepth + 1 : 1;
            FullPath = fullPath?.FromVbaStringLiteral() ?? string.Empty;
            Name = name.Replace("\"", string.Empty);

            AddNewChildren(ref declarations);
        }

        public override string Name { get; }

        public override string PanelTitle => FullPath ?? string.Empty;

        public override string Description => FolderAttribute ?? string.Empty;

        public string FullPath { get; }

        public string FolderAttribute => $"'@Folder({FullPath.ToVbaStringLiteral()})";

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

        protected override void AddNewChildren(ref List<Declaration> declarations)
        {
            var children = declarations.Where(declaration => declaration.IsInFolderOrSubFolder(FullPath)).ToList();
            declarations = declarations.Except(children).ToList();

            var subFolders = children.Where(declaration => declaration.IsInSubFolder(FullPath)).ToList();

            foreach (var folder in subFolders.GroupBy(declaration => declaration.CustomFolder.SubFolderRoot(FullPath)))
            {
                var contents = folder.ToList();
                AddChild(new CodeExplorerCustomFolderViewModel(this, folder.Key, $"{FullPath}.{folder.Key}", _vbe, ref contents));
            }

            children = children.Except(subFolders).ToList();

            foreach (var declaration in children.Where(child => child.IsInFolder(FullPath)).GroupBy(item => item.ComponentName))
            {
                var moduleName = declaration.Key;
                var parent = children.SingleOrDefault(item => 
                    ComponentTypes.Contains(item.DeclarationType) && item.ComponentName == moduleName);

                if (parent is null)
                {
                    continue;
                }

                var members = children.Where(item =>
                    !ComponentTypes.Contains(item.DeclarationType) && item.ComponentName == moduleName).ToList();

                AddChild(new CodeExplorerComponentViewModel(this, parent, ref members, _vbe));
            }
        }

        public override void Synchronize(ref List<Declaration> updated)
        {
            SynchronizeChildren(ref updated);
        }

        protected override void SynchronizeChildren(ref List<Declaration> updated)
        {
            var children = updated.Where(declaration => declaration.IsInFolderOrSubFolder(FullPath)).ToList();
            updated = updated.Except(children).ToList();

            if (!children.Any())
            {
                Declaration = null;
                return;
            }

            var subFolders = children.Where(declaration => declaration.IsInSubFolder(FullPath)).ToList();
            children = children.Except(subFolders).ToList();

            foreach (var subfolder in Children.OfType<CodeExplorerCustomFolderViewModel>().ToList())
            {
                subfolder.SynchronizeChildren(ref subFolders);
                if (subfolder.Declaration is null)
                {
                    RemoveChild(subfolder);
                }
            }

            foreach (var child in Children.OfType<CodeExplorerComponentViewModel>().ToList())
            {
                child.Synchronize(ref children);
                if (child.Declaration is null)
                {
                    RemoveChild(child);
                }
            }

            children = children.Concat(subFolders).ToList();
            AddNewChildren(ref children);
        }
    }
}
