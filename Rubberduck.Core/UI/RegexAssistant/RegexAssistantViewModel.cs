using Rubberduck.RegexAssistant;
using System;
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

        public bool GlobalFlag
        {
            get => _globalFlag;
            set
            {
                _globalFlag = value;
                RecalculateDescription();
            }
        }

        public bool IgnoreCaseFlag
        {
            get => _ignoreCaseFlag;
            set
            {
                _ignoreCaseFlag = value;
                RecalculateDescription();
            }
        }

        public string Pattern
        {
            get => _pattern;
            set
            {
                _pattern = value;
                RecalculateDescription();
            }
        }

        private bool _globalFlag;
        private bool _ignoreCaseFlag;
        private string _pattern;

        private List<TreeViewItem> _resultItems;
        public List<TreeViewItem> ResultItems
        {
            get => _resultItems;
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
                DescriptionResults = RubberduckUI.RegexAssistant_NoPatternError;
                var results = new List<TreeViewItem>();

                var rootTreeItem = new TreeViewItem
                {
                    Header = DescriptionResults
                };

                results.Add(rootTreeItem);
                ResultItems = results;
                return;
            }
            ResultItems = ToTreeViewItems(new Pattern(_pattern, _ignoreCaseFlag, _globalFlag));
        }

        private List<TreeViewItem> ToTreeViewItems(Pattern pattern)
        {
            var resultItems = new List<TreeViewItem>();
            if (pattern.IgnoreCase)
            {
                resultItems.Add(TreeViewItemFromHeader(pattern.CasingDescription));
            }
            if (pattern.AnchoredAtStart)
            {
                resultItems.Add(TreeViewItemFromHeader(pattern.StartAnchorDescription));
            }
            resultItems.Add(AsTreeViewItem((dynamic)pattern.RootExpression));
            if (pattern.AnchoredAtEnd)
            {
                resultItems.Add(TreeViewItemFromHeader(pattern.EndAnchorDescription));
            }
            return resultItems;
        }

        private TreeViewItem TreeViewItemFromHeader(string header)
        {
            var result = new TreeViewItem
            {
                Header = header
            };

            return result;
        }

        public string DescriptionResults { get; private set; }

        private static TreeViewItem AsTreeViewItem(IRegularExpression expression)
        {
            throw new InvalidOperationException("Some unknown IRegularExpression subtype was in RegexAssistantViewModel");
        }

        private static TreeViewItem AsTreeViewItem(ErrorExpression expression)
        {
            var result = new TreeViewItem
            {
                Header = expression.Description
            };

            return result;
        }

        private static TreeViewItem AsTreeViewItem(ConcatenatedExpression expression)
        {
            var result = new TreeViewItem
            {
                Header = expression.Description
            };

            foreach (var subtree in expression.Subexpressions.Select(exp => AsTreeViewItem((dynamic)exp)))
            {
                result.Items.Add(subtree);
            }
            return result;
        }

        private static TreeViewItem AsTreeViewItem(AlternativesExpression expression)
        {
            var result = new TreeViewItem
            {
                Header = expression.Description
            };

            foreach (var subtree in expression.Subexpressions.Select(exp => AsTreeViewItem((dynamic)exp)))
            {
                result.Items.Add(subtree);
            }
            return result;
        }

        private static TreeViewItem AsTreeViewItem(SingleAtomExpression expression)
        {
            var result = new TreeViewItem
            {
                Header = expression.Description
            };

            // no other Atom has Subexpressions we care about
            if (expression.Atom.GetType() == typeof(Group))
            {
                result.Items.Add(AsTreeViewItem((dynamic)((expression.Atom) as Group).Subexpression));
            }
            
            return result;
        }
    }
}
