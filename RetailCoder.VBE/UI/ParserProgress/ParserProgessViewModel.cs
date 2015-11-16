using System.Windows.Media.Imaging;
using Microsoft.Vbe.Interop;
using Rubberduck.Common;

namespace Rubberduck.UI.ParserProgress
{
    public class ComponentProgressViewModel : ViewModelBase
    {
        private readonly VBComponent _component;

        public ComponentProgressViewModel(VBComponent component)
        {
            _component = component;
            ComponentName = component.Name;
        }

        public BitmapImage ComponentIcon
        {
            get { return DeclarationIconCache.ComponentIcon(_component.Type); }
        }

        public string ComponentName { get; private set; }

        private decimal _value;

        public decimal ResolutionProgressPercent
        {
            get { return _value; }
            set
            {
                _value = value;
                OnPropertyChanged();
            }
        }
    }
}
