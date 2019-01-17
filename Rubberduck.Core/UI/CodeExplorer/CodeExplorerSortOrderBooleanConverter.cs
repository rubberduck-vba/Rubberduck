using System;
using System.Globalization;
using System.Windows.Data;
using Rubberduck.Navigation.CodeExplorer;

namespace Rubberduck.UI.CodeExplorer
{
    /// <summary>
    /// Binds an individual flag of CodeExplorerSortOrder to a boolean. Note: This is a stateful converter, so each bound control
    /// requires its own converter instance.
    /// </summary>
    internal class CodeExplorerSortOrderBooleanConverter : IValueConverter
    {
        private CodeExplorerSortOrder _state;

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (!(parameter is CodeExplorerSortOrder flag) ||
                !(value is CodeExplorerSortOrder bound))
            {
                return _state;
            }

            _state = bound;
            return _state.HasFlag(flag);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (!(parameter is CodeExplorerSortOrder flag) ||
                !(value is bool isSet))
            {
                return _state;
            }

            switch (flag)
            {
                case CodeExplorerSortOrder.Name:
                    _state = isSet ? (_state & ~CodeExplorerSortOrder.CodeLine) | flag : (_state & ~flag) | CodeExplorerSortOrder.Name;
                    break;
                case CodeExplorerSortOrder.CodeLine:
                    _state = isSet ? (_state & ~CodeExplorerSortOrder.Name) | flag : (_state & ~flag) | CodeExplorerSortOrder.CodeLine;
                    break;
                default:
                    _state ^= flag;
                    break;
            }

            return _state;
        }
    }
}
