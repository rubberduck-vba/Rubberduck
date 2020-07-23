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
            //Note: this is necessary for a hack dealing with the fact that a named element on the control cannot be accessed from the context menu of the element by name.
            var selectedItems = values[1] as IEnumerable
                                ?? values[2] as IEnumerable
                                ?? new List<IInspectionResult>();

            (IQuickFix quickFix, IEnumerable<IInspectionResult> selectedResults) data = (
                values[0] as IQuickFix,
                selectedItems.OfType<IInspectionResult>());

            return data;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}