using System.Collections.Generic;
using System.Linq;
using System.Windows.Media.Imaging;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Navigation.CodeExplorer
{
    public class CompareByName : Comparer<CodeExplorerItemViewModel>
    {
        public override int Compare(CodeExplorerItemViewModel x, CodeExplorerItemViewModel y)
        {
            if (x == y)
            {
                return 0;
            }

            var nodeComparison = new CompareByNodeType().Compare(x, y);
            if (nodeComparison != 0)
            {
                return nodeComparison;
            }

            return string.CompareOrdinal(x.NameWithSignature, y.NameWithSignature);
        }
    }

    public class CompareByType : Comparer<CodeExplorerItemViewModel>
    {
        private static Dictionary<DeclarationType, int> _sortOrder = new Dictionary<DeclarationType, int>
        {
            {DeclarationType.LibraryFunction, 0},
            {DeclarationType.LibraryProcedure, 1},
            {DeclarationType.UserDefinedType, 2},
            {DeclarationType.Enumeration, 3},
            {DeclarationType.Event, 4},
            {DeclarationType.Variable, 5},
            {DeclarationType.PropertyGet, 6},
            {DeclarationType.PropertyLet, 7},
            {DeclarationType.PropertySet, 8},
            {DeclarationType.Function, 9},
            {DeclarationType.Procedure, 10}
        };

        public override int Compare(CodeExplorerItemViewModel x, CodeExplorerItemViewModel y)
        {
            if (x == y)
            {
                return 0;
            }

            var nodeComparison = new CompareByNodeType().Compare(x, y);
            if (nodeComparison != 0)
            {
                return nodeComparison;
            }

            // error nodes have the same sort value
            if (x is CodeExplorerErrorNodeViewModel &&
                y is CodeExplorerErrorNodeViewModel)
            {
                return 0;
            }

            var xNode = (ICodeExplorerDeclarationViewModel)x;
            var yNode = (ICodeExplorerDeclarationViewModel)y;

            // keep separate types separate
            if (xNode.Declaration.DeclarationType != yNode.Declaration.DeclarationType)
            {
                return _sortOrder[xNode.Declaration.DeclarationType] < _sortOrder[yNode.Declaration.DeclarationType] ? -1 : 1;
            }

            // keep types with different icons and the same declaration type (document/class module) separate
            // documents come first
            if (x.ExpandedIcon != y.ExpandedIcon)
            {
                // ReSharper disable once PossibleInvalidOperationException - this will have a component
                return x.QualifiedSelection.Value.QualifiedName.Component.Type ==
                       Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_Document
                    ? -1
                    : 1;
            }

            return 0;
        }
    }

    public class CompareBySelection : Comparer<CodeExplorerItemViewModel>
    {
        public override int Compare(CodeExplorerItemViewModel x, CodeExplorerItemViewModel y)
        {
            if (x == y)
            {
                return 0;
            }

            var nodeComparison = new CompareByNodeType().Compare(x, y);
            if (nodeComparison != 0)
            {
                return nodeComparison;
            }

            if (!x.QualifiedSelection.HasValue && !y.QualifiedSelection.HasValue)
            {
                return 0;
            }

            if (x.QualifiedSelection.HasValue ^ y.QualifiedSelection.HasValue)
            {
                return x.QualifiedSelection.HasValue ? -1 : 1;
            }

            if (x.QualifiedSelection.Value.Selection == y.QualifiedSelection.Value.Selection)
            {
                return 0;
            }

            return x.QualifiedSelection.Value.Selection < y.QualifiedSelection.Value.Selection ? -1 : 1;
        }
    }

    public class CompareByNodeType : Comparer<CodeExplorerItemViewModel>
    {
        public override int Compare(CodeExplorerItemViewModel x, CodeExplorerItemViewModel y)
        {
            if (x == y)
            {
                return 0;
            }

            // folders come first
            if (x is CodeExplorerCustomFolderViewModel ^
                y is CodeExplorerCustomFolderViewModel)
            {
                return x is CodeExplorerCustomFolderViewModel ? -1 : 1;
            }

            // folders are always sorted by name
            if (x is CodeExplorerCustomFolderViewModel &&
                y is CodeExplorerCustomFolderViewModel)
            {
                return string.CompareOrdinal(x.NameWithSignature, y.NameWithSignature);
            }

            // error nodes come after folders
            if (x is CodeExplorerErrorNodeViewModel ^
                y is CodeExplorerErrorNodeViewModel)
            {
                return x is CodeExplorerErrorNodeViewModel ? -1 : 1;
            }

            return 0;
        }
    }

    public abstract class CodeExplorerItemViewModel : ViewModelBase
    {
        private List<CodeExplorerItemViewModel> _items = new List<CodeExplorerItemViewModel>();
        public List<CodeExplorerItemViewModel> Items
        {
            get { return _items; }
            protected set
            {
                _items = value;
                OnPropertyChanged();
            }
        }

        public bool IsExpanded { get; set; }

        public abstract string Name { get; }
        public abstract string NameWithSignature { get; }
        public abstract BitmapImage CollapsedIcon { get; }
        public abstract BitmapImage ExpandedIcon { get; }
        public abstract CodeExplorerItemViewModel Parent { get; }

        public abstract QualifiedSelection? QualifiedSelection { get; }

        public CodeExplorerItemViewModel GetChild(string name)
        {
            foreach (var item in _items)
            {
                if (item.Name == name)
                {
                    return item;
                }
                var result = item.GetChild(name);
                if (result != null)
                {
                    return result;
                }
            }

            return null;
        }

        public Declaration GetSelectedDeclaration()
        {
            return this is ICodeExplorerDeclarationViewModel
                ? ((ICodeExplorerDeclarationViewModel)this).Declaration
                : null;
        }

        public void AddChild(CodeExplorerItemViewModel item)
        {
            _items.Add(item);
        }

        public void ReorderItems(bool sortByName, bool sortByType)
        {
            if (sortByType)
            {
                Items = sortByName
                    ? Items.OrderBy(o => o, new CompareByType()).ThenBy(t => t, new CompareByName()).ToList()
                    : Items.OrderBy(o => o, new CompareByType()).ThenBy(t => t, new CompareBySelection()).ToList();

                return;
            }

            Items = sortByName
                ? Items.OrderBy(t => t, new CompareByName()).ToList()
                : Items.OrderBy(t => t, new CompareBySelection()).ToList();
        }
    }
}
