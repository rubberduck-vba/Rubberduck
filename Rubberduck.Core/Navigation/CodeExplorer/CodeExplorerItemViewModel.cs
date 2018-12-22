using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Media.Imaging;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Navigation.CodeExplorer
{
    public abstract class CodeExplorerItemViewModel : ViewModelBase
    {
        protected CodeExplorerItemViewModel(Declaration declaration)
        {
            Declaration = declaration;
        }

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

        private bool _selected;
        public bool IsSelected
        {
            get => _selected;
            set
            {
                _selected = value;
                OnPropertyChanged();
            }
        }

        private bool _dimmed;
        public virtual bool IsDimmed
        {
            get => _dimmed;
            set
            {
                _dimmed = value;
                OnPropertyChanged();
            }
        }

        private bool _isVisible = true;
        public bool IsVisible
        {
            get => _isVisible;
            set
            {
                _isVisible = value;
                OnPropertyChanged();
            }
        }

        public virtual bool IsObsolete => false;

        public Declaration Declaration { get; }

        public virtual string ToolTip => NameWithSignature;

        public abstract string Name { get; }
        public abstract string NameWithSignature { get; }
        public abstract BitmapImage CollapsedIcon { get; }
        public abstract BitmapImage ExpandedIcon { get; }
        public abstract CodeExplorerItemViewModel Parent { get; }

        public virtual FontWeight FontWeight => FontWeights.Normal;

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
