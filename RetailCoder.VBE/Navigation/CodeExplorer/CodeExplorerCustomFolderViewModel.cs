using System.Collections.Generic;
using System.Linq;
using System.Windows.Media.Imaging;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using resx = Rubberduck.Properties.Resources;

namespace Rubberduck.Navigation.CodeExplorer
{
    public class CodeExplorerCustomFolderViewModel : ViewModelBase
    {
        private readonly string _name;
        private readonly IEnumerable<CodeExplorerComponentViewModel> _components;

        private static readonly DeclarationType[] ComponentTypes =
        {
            DeclarationType.Class, 
            DeclarationType.Document, 
            DeclarationType.Module, 
            DeclarationType.UserForm, 
        };

        public CodeExplorerCustomFolderViewModel(string name, IEnumerable<Declaration> declarations)
        {
            _name = name;

            var items = declarations.ToList();

            _components = items.GroupBy(item => item.ComponentName)
                .SelectMany(grouping =>
                    grouping.Where(item => ComponentTypes.Contains(item.DeclarationType))
                        .Select(item => new CodeExplorerComponentViewModel(item, grouping)))
                .OrderBy(item => item.Name)
                .ToList();

            _blueFolderCollapsed = GetImageSource(resx.blue_folder_horizontal);
            _blueFolderExpanded = GetImageSource(resx.blue_folder_horizontal_open);
        }

        private readonly BitmapImage _blueFolderCollapsed;
        public BitmapImage BlueFolderCollapsed { get { return _blueFolderCollapsed; } }

        private readonly BitmapImage _blueFolderExpanded;
        public BitmapImage BlueFolderExpanded { get { return _blueFolderExpanded; } }


        public string Name { get { return _name; } }

        public IEnumerable<CodeExplorerComponentViewModel> Components { get { return _components; } }
    }
}