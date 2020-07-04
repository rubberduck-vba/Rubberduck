using System.Windows;

namespace Rubberduck.UI.Inspections
{
    /// <summary>
    /// Interaction logic for InspectionResultsControl.xaml
    /// </summary>
    public partial class InspectionResultsControl
    {
        private const int HorizontalRectangleAdjustment = 2000;

        private InspectionResultsViewModel ViewModel => DataContext as InspectionResultsViewModel;

        public InspectionResultsControl()
        {
            InitializeComponent();
        }

        //Based on https://stackoverflow.com/a/42238409/5536802 by Jason Williams and the comment to it by Nick Desjardins.
        private bool _requestingModifiedBringIntoView;
        private void InspectionResultsGrid_RequestBringIntoView(object sender, RequestBringIntoViewEventArgs e)
        {
            if (_requestingModifiedBringIntoView
                || !(e?.OriginalSource is FrameworkElement source))
            {
                return;
            }

            e.Handled = true;

            //Prevents adjustment of the adjusted event triggered below.
            _requestingModifiedBringIntoView = true;

            var newRectangle = new Rect(-HorizontalRectangleAdjustment, 0, source.ActualWidth + HorizontalRectangleAdjustment, source.ActualHeight);
            source.BringIntoView(newRectangle);

            _requestingModifiedBringIntoView = false;
        }
    }
}
