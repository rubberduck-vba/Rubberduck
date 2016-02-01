using System.Collections.Generic;
using System.Linq;
using System.Windows.Media.Imaging;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Navigation.CodeExplorer
{
    public abstract class CodeExplorerItemViewModel : ViewModelBase
    {
        private IList<CodeExplorerItemViewModel> _items = new List<CodeExplorerItemViewModel>();
        public IEnumerable<CodeExplorerItemViewModel> Items { get { return _items; } protected set { _items = value.ToList(); } }

        public abstract string Name { get; }
        public abstract BitmapImage CollapsedIcon { get; }
        public abstract BitmapImage ExpandedIcon { get; }

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
    }
}