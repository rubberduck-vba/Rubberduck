using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Data;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.QuickFixes;

namespace Rubberduck.UI.Inspections
{
    public class QuickFixCommandParametersToTupleConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            (IQuickFix quickFix, IEnumerable<IInspectionResult> selectedResults) data = (
                values[0] as IQuickFix,
                ((IEnumerable)values[1]).OfType<IInspectionResult>());
            return data;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}