using Rubberduck.RegexAssistant;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Controls;

namespace Rubberduck.UI.RegexAssistant
{
    public class RegexAssistantViewModel : ViewModelBase
    {
        public RegexAssistantViewModel()
        {
            _pattern = string.Empty;
            RecalculateDescription();
        }

        public bool GlobalFlag {
            get
            {
                return _globalFlag;
            }
            set
            {
                _globalFlag = value;
                RecalculateDescription();
            }
        }
        public bool IgnoreCaseFlag
        {
            get
            {
                return _ignoreCaseFlag;
            }
            set
            {
                _ignoreCaseFlag = value;
                RecalculateDescription();
            }
        }
        public string Pattern
        {
            get
            {
                return _pattern;
            }
            set
            {
                _pattern = value;
                RecalculateDescription();
            }
        }

        private string _description;
        private bool _globalFlag;
        private bool _ignoreCaseFlag;
        private string _pattern;

        private List<TreeViewItem> _resultItems;
        public List<TreeViewItem> ResultItems
        {
            get
            {
                return _resultItems;
            }
            set
            {
                _resultItems = value;
                OnPropertyChanged();
            }
        }

        private void RecalculateDescription()
        {
            if (_pattern.Equals(string.Empty))
            {
                _description = "No Pattern given";
                var results = new List<TreeViewItem>();
                var rootTreeItem = new TreeViewItem();
                rootTreeItem.Header = _description;
                results.Add(rootTreeItem);
                ResultItems = results;
                return;
            }
            var pattern = new Pattern(_pattern, _ignoreCaseFlag, _ignoreCaseFlag);
            //_description = pattern.Description;
            var resultItems = new List<TreeViewItem>();
            resultItems.Add(AsTreeViewItem((dynamic)pattern.RootExpression));
            ResultItems = resultItems;
            //base.OnPropertyChanged("DescriptionResults");
        }

        public string DescriptionResults
        {
            get
            {
                return _description;
            }
        }

        private static TreeViewItem AsTreeViewItem(IRegularExpression expression)
        {
            var result = new TreeViewItem();
            result.Header = "Some unknown IRegularExpression subtype was in the view";
            foreach (var subtree in expression.Subexpressions.Select(exp => AsTreeViewItem((dynamic)exp)))
            {
                result.Items.Add(subtree);
            }
            return result;
        }

        private static TreeViewItem AsTreeViewItem(ErrorExpression expression)
        {
            var result = new TreeViewItem();
            result.Header = expression.Description;
            return result;
        }

        private static TreeViewItem AsTreeViewItem(ConcatenatedExpression expression)
        {
            var result = new TreeViewItem();
            result.Header = expression.Description;
            foreach (var subtree in expression.Subexpressions.Select(exp => AsTreeViewItem((dynamic)exp)))
            {
                result.Items.Add(subtree);
            }
            return result;
        }

        private static TreeViewItem AsTreeViewItem(AlternativesExpression expression)
        {
            var result = new TreeViewItem();
            result.Header = expression.Description;
            foreach (var subtree in expression.Subexpressions.Select(exp => AsTreeViewItem((dynamic)exp)))
            {
                result.Items.Add(subtree);
            }
            return result;
        }

        private static TreeViewItem AsTreeViewItem(SingleAtomExpression expression)
        {
            var result = new TreeViewItem();
            result.Header = expression.Description;
            // no other Atom has Subexpressions we care about
            if (expression.Atom.GetType() == typeof(Group))
            {
                result.Items.Add(AsTreeViewItem((dynamic)((expression.Atom) as Group).Subexpression));
            }
            
            return result;
        }
    }
}
