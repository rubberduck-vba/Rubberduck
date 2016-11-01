using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Media.Imaging;
using Rubberduck.Navigation.Folders;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using resx = Rubberduck.UI.CodeExplorer.CodeExplorer;

namespace Rubberduck.Navigation.CodeExplorer
{
    public class CodeExplorerProjectViewModel : CodeExplorerItemViewModel, ICodeExplorerDeclarationViewModel
    {
        private readonly Declaration _declaration;
        public Declaration Declaration { get { return _declaration; } }
        private readonly CodeExplorerCustomFolderViewModel _folderTree;

        private static readonly DeclarationType[] ComponentTypes =
        {
            DeclarationType.ClassModule, 
            DeclarationType.Document, 
            DeclarationType.ProceduralModule, 
            DeclarationType.UserForm, 
        };

        public CodeExplorerProjectViewModel(FolderHelper folderHelper, Declaration declaration, IEnumerable<Declaration> declarations)
        {
            _declaration = declaration;
            _name = _declaration.IdentifierName;
            IsExpanded = true;
            _folderTree = folderHelper.GetFolderTree(declaration);

            try
            {
                FillFolders(declarations.ToList());
                Items = _folderTree.Items.ToList();

                _icon = _declaration.Project.Protection == ProjectProtection.Locked
                    ? GetImageSource(resx.lock__exclamation)
                    : GetImageSource(resx.ObjectLibrary);
            }
            catch (NullReferenceException e)
            {
                Console.WriteLine(e);
            }
        }

        private void FillFolders(IEnumerable<Declaration> declarations)
        {
            var items = declarations.ToList();
            var groupedItems = items.Where(item => ComponentTypes.Contains(item.DeclarationType))
                               .GroupBy(item => item.CustomFolder)
                               .OrderBy(item => item.Key);

            // set parent so we can walk up to the project node
            // we haven't added the nodes yet, so this cast is valid
            // ReSharper disable once PossibleInvalidCastExceptionInForeachLoop
            foreach (CodeExplorerCustomFolderViewModel item in _folderTree.Items)
            {
                item.SetParent(this);
            }

            foreach (var grouping in groupedItems)
            {
                AddNodesToTree(_folderTree, items, grouping);
            }
        }

        private bool AddNodesToTree(CodeExplorerCustomFolderViewModel tree, List<Declaration> items, IGrouping<string, Declaration> grouping)
        {
            foreach (var folder in tree.Items.OfType<CodeExplorerCustomFolderViewModel>())
            {
                if (grouping.Key.Replace("\"", string.Empty) != folder.FullPath)
                {
                    continue;
                }

                var parents = grouping.Where(
                        item => ComponentTypes.Contains(item.DeclarationType) &&
                            item.CustomFolder.Replace("\"", string.Empty) == folder.FullPath)
                        .ToList();

                folder.AddNodes(items.Where(item => parents.Contains(item) || parents.Any(parent =>
                    (item.ParentDeclaration != null && item.ParentDeclaration.Equals(parent)) ||
                    item.ComponentName == parent.ComponentName)).ToList());

                return true;
            }

            return tree.Items.OfType<CodeExplorerCustomFolderViewModel>().Any(node => AddNodesToTree(node, items, grouping));
        }

        private readonly BitmapImage _icon;
        public override BitmapImage CollapsedIcon { get { return _icon; } }
        public override BitmapImage ExpandedIcon { get { return _icon; } }
        
        // projects are always at the top of the tree
        public override CodeExplorerItemViewModel Parent { get { return null; } }

        private string _name;
        public override string Name { get { return _name; } }
        public override string NameWithSignature { get { return _name; } }
        public override QualifiedSelection? QualifiedSelection { get { return _declaration.QualifiedSelection; } }

        public void SetParenthesizedName(string parenthesizedName)
        {
            _name += " (" + parenthesizedName + ")";
        }
    }
}
