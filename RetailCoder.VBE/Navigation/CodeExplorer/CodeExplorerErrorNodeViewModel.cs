using System.Windows.Media.Imaging;
using Rubberduck.VBEditor;
using resx = Rubberduck.UI.CodeExplorer.CodeExplorer;

namespace Rubberduck.Navigation.CodeExplorer
{
    public class CodeExplorerErrorNodeViewModel : CodeExplorerItemViewModel
    {
        public CodeExplorerErrorNodeViewModel(string name)
        {
            _name = name;
        }

        private readonly string _name;
        public override string Name { get { return _name; } }
        public override string NameWithSignature { get { return _name; } }

        public override BitmapImage CollapsedIcon { get { return GetImageSource(resx.Error); } }
        public override BitmapImage ExpandedIcon { get { return GetImageSource(resx.Error); } }

        public override QualifiedSelection? QualifiedSelection { get { return null; } }
    }
}