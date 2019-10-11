using Rubberduck.UI.UnitTesting.ViewModels;
using Rubberduck.Common;

namespace Rubberduck.Formatters
{
    public class TestMethodViewModelFormatter : IExportable
    {
        private readonly TestMethodViewModel _testMethodViewModel;

        public TestMethodViewModelFormatter(TestMethodViewModel testMethodViewModel)
        {
            _testMethodViewModel = testMethodViewModel;
        }

        public object[] ToArray()
        {
            return _testMethodViewModel.ToArray();
        }

        public string ToClipboardString()
        {
            return _testMethodViewModel.ToString();
        }
    }
}
