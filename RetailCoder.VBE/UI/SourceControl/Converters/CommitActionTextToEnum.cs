using System;
using System.Globalization;
using System.Windows.Data;

namespace Rubberduck.UI.SourceControl.Converters
{
    public class CommitActionTextToEnum : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var action = (CommitAction)value;
            switch (action)
            {
                case CommitAction.Commit:
                    return RubberduckUI.SourceControl_Commit;
                case CommitAction.CommitAndPush:
                    return RubberduckUI.SourceControl_CommitAndPush;
                case CommitAction.CommitAndSync:
                    return RubberduckUI.SourceControl_CommitAndSync;
                default:
                    return value;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var text = (string)value;

            if (text == RubberduckUI.SourceControl_Commit)
            {
                return CommitAction.Commit;
            }
            if (text == RubberduckUI.SourceControl_CommitAndPush)
            {
                return CommitAction.CommitAndPush;
            }
            if (text == RubberduckUI.SourceControl_CommitAndSync)
            {
                return CommitAction.CommitAndSync;
            }

            return null;
        }
    }
}
