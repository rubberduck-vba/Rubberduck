using System.Collections.Generic;
using System.Linq;
using Rubberduck.Navigation.Folders;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Navigation.CodeExplorer
{
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

            FullPath = fullPath ?? string.Empty;
            Name = name.Replace("\"", string.Empty);

            AddNewChildren(declarations.ToList());
        }

        public override string Name { get; }

        public override string PanelTitle => FullPath ?? string.Empty;

        public override string Description => FolderAttribute ?? string.Empty;

        public string FullPath { get; }

        public string FolderAttribute => $"@Folder(\"{FullPath.Replace("\"", string.Empty)}\")";

        public override QualifiedSelection? QualifiedSelection => null;

        public override bool IsErrorState
        {
            get => false;
            set { }
        }

        public override Comparer<ICodeExplorerNode> SortComparer => CodeExplorerItemComparer.Name;

        public override bool Filtered => false;

        protected override void AddNewChildren(List<Declaration> declarations)
        {
            var subfolders = declarations.Where(declaration => declaration.IsInSubFolder(FullPath)).ToList();

            foreach (var folder in subfolders.GroupBy(declaration => declaration.CustomFolder))
            {
                AddChild(new CodeExplorerCustomFolderViewModel(this, folder.Key.SubFolderRoot(Name), folder.Key, _vbe, folder));
            }

            var components = declarations.Except(subfolders).ToList();

            foreach (var component in components.GroupBy(item => item.ComponentName))
            {
                var moduleName = component.Key;
                var parent = declarations.SingleOrDefault(item => 
                    ComponentTypes.Contains(item.DeclarationType) && item.ComponentName == moduleName);

                if (parent is null)
                {
                    continue;
                }

                var members = declarations.Where(item =>
                    !ComponentTypes.Contains(item.DeclarationType) && item.ComponentName == moduleName);

                AddChild(new CodeExplorerComponentViewModel(this, parent, members, _vbe));
            }
        }

        public override void Synchronize(List<Declaration> updated)
        {
            var declarations = updated.Where(declaration => declaration.IsInFolderOrSubFolder(FullPath)).ToList();

            if (!declarations.Any())
            {
                Declaration = null;
                return;
            }

            foreach (var declaration in declarations)
            {
                updated.Remove(declaration);
            }

            SynchronizeChildren(declarations);
        }
    }
}
