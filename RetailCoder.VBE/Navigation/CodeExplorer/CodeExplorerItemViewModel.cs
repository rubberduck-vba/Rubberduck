using System.Collections.Generic;
using System.Linq;
using System.Windows.Media.Imaging;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;

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

            return nodeComparison != 0
                ? nodeComparison
                : string.CompareOrdinal(x.NameWithSignature, y.NameWithSignature);
        }
    }

    public class CompareByType : Comparer<CodeExplorerItemViewModel>
    {
        private static readonly Dictionary<DeclarationType, int> SortOrder = new Dictionary<DeclarationType, int>
        {
            // Some DeclarationTypes we want to treat the same, like Subs and Functions,
            // or Property Gets, Lets, and Sets.
            // Give them the same number.
            {DeclarationType.LibraryFunction, 0},
            {DeclarationType.LibraryProcedure, 0},
            {DeclarationType.UserDefinedType, 1},
            {DeclarationType.Enumeration, 2},
            {DeclarationType.Event, 3},
            {DeclarationType.Constant, 4},
            {DeclarationType.Variable, 5},
            {DeclarationType.PropertyGet, 6},
            {DeclarationType.PropertyLet, 6},
            {DeclarationType.PropertySet, 6},
            {DeclarationType.Function, 7},
            {DeclarationType.Procedure, 7}
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

            var xNode = (ICodeExplorerDeclarationViewModel)x;
            var yNode = (ICodeExplorerDeclarationViewModel)y;

            // keep separate types separate
            if (xNode.Declaration.DeclarationType != yNode.Declaration.DeclarationType)
            {
                if (SortOrder.TryGetValue(xNode.Declaration.DeclarationType, out var xValue) &&
                    SortOrder.TryGetValue(yNode.Declaration.DeclarationType, out var yValue))
                {
                    if (xValue != yValue)
                    { return xValue < yValue ? -1 : 1; }
                }
            }

            // The Tree shows Public and Private Subs/Functions with a seperate icon.
            // But Public and Implicit Subs/Functions appear the same, so treat Implicts like Publics.
            var xNodeAcc = xNode.Declaration.Accessibility == Accessibility.Implicit ? Accessibility.Public : xNode.Declaration.Accessibility;
            var yNodeAcc = yNode.Declaration.Accessibility == Accessibility.Implicit ? Accessibility.Public : yNode.Declaration.Accessibility;

            if (xNodeAcc != yNodeAcc)
            {
                return xNodeAcc > yNodeAcc ? -1 : 1;
            }

            if (x.ExpandedIcon != y.ExpandedIcon)
            {
                // ReSharper disable PossibleInvalidOperationException - this will have a component
                var xComponent = x.QualifiedSelection.Value.QualifiedName.Component;
                var yComponent = y.QualifiedSelection.Value.QualifiedName.Component;

                if (xComponent.Type == ComponentType.Document ^ yComponent.Type == ComponentType.Document)
                {
                    return xComponent.Type == ComponentType.Document ? -1 : 1;
                }
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

            if (x.QualifiedSelection.Value.Selection.StartLine == y.QualifiedSelection.Value.Selection.StartLine)
            {
                return 0;
            }

            return x.QualifiedSelection.Value.Selection.StartLine < y.QualifiedSelection.Value.Selection.StartLine ? -1 : 1;
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

            return 0;
        }
    }

    public abstract class CodeExplorerItemViewModel : ViewModelBase
    {
        private List<CodeExplorerItemViewModel> _items = new List<CodeExplorerItemViewModel>();
        public List<CodeExplorerItemViewModel> Items
        {
            get => _items;
            protected set
            {
                _items = value;
                OnPropertyChanged();
            }
        }

        private bool _isExpanded;
        public bool IsExpanded
        {
            get => _isExpanded;
            set
            {
                _isExpanded = value;
                OnPropertyChanged();
            }
        }

        public bool IsSelected { get; set; }

        private bool _isVisisble = true;
        public bool IsVisible
        {
            get => _isVisisble;
            set
            {
                _isVisisble = value;
                OnPropertyChanged();
            }
        }

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
            return this is ICodeExplorerDeclarationViewModel viewModel
                ? viewModel.Declaration
                : null;
        }

        public void AddChild(CodeExplorerItemViewModel item)
        {
            _items.Add(item);
        }

        public void ReorderItems(bool sortByName, bool groupByType)
        {
            if (groupByType)
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
