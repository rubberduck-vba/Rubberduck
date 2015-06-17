using System;
using System.Linq;

namespace Rubberduck.UI.Settings
{
    public class AddMarkerFormPresenter
    {
        private readonly IAddTodoMarkerView _view;

        public AddMarkerFormPresenter(IAddTodoMarkerView view)
        {
            _view = view;

            _view.TextChanged += TextChanged;
        }

        private void TextChanged(object sender, EventArgs e)
        {
            _view.IsValidMarker = _view.TodoMarkers.All(t => t.Text != _view.MarkerText) && _view.MarkerText != string.Empty;
        }
    }
}
