using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;
using Rubberduck.UnitTesting;
using Rubberduck.Resources;
using Rubberduck.UI.UnitTesting.ViewModels;
using ImageSourceConverter = Rubberduck.UI.Converters.ImageSourceConverter;

namespace Rubberduck.UI.UnitTesting
{
    public class TestOutcomeImageSourceConverter : ImageSourceConverter, IMultiValueConverter
    {
        private static readonly ImageSource QueuedIcon = ToImageSource(RubberduckUI.clock);
        private static readonly ImageSource RunningIcon = ToImageSource(RubberduckUI.hourglass);

        private static readonly IDictionary<TestOutcome,ImageSource> Icons = 
            new Dictionary<TestOutcome, ImageSource>
            {
                { TestOutcome.Unknown, ToImageSource(RubberduckUI.question_white) },
                { TestOutcome.Succeeded, ToImageSource(RubberduckUI.tick_circle) },
                { TestOutcome.Failed, ToImageSource(RubberduckUI.cross_circle) },
                { TestOutcome.Inconclusive, ToImageSource(RubberduckUI.exclamation) },
                { TestOutcome.Ignored, ToImageSource(RubberduckUI.minus_white) }
            };

        public override object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (!(value is TestMethodViewModel test))
            {
                return null;
            }

            if (test.RunState != TestRunState.Stopped)
            {
                return test.RunState == TestRunState.Running ? RunningIcon : QueuedIcon;
            }

            var outcome = test.Result.Outcome;
            return Icons[outcome];
        }

        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            return values.Length == 0 ? null : Convert(values[0], targetType, parameter, culture);
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
