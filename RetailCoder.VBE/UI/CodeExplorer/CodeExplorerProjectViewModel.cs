using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Media.Imaging;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;
using resx = Rubberduck.Properties.Resources;

namespace Rubberduck.UI.CodeExplorer
{
    public class CodeExplorerProjectViewModel : ViewModelBase
    {
        private readonly Declaration _declaration;
        private readonly IEnumerable<CodeExplorerComponentViewModel> _components;
        private readonly IEnumerable<CodeExplorerCustomFolderViewModel> _customFolders; 

        private readonly BitmapImage _icon;

        private static readonly DeclarationType[] ComponentTypes =
        {
            DeclarationType.Class, 
            DeclarationType.Document, 
            DeclarationType.Module, 
            DeclarationType.UserForm, 
        };

        public CodeExplorerProjectViewModel(Declaration declaration, IEnumerable<Declaration> declarations)
        {
            _declaration = declaration;
            var items = declarations.ToList();

            _components = items.GroupBy(item => item.ComponentName)
                .SelectMany(grouping =>
                    grouping.Where(item => ComponentTypes.Contains(item.DeclarationType))
                        .Select(item => new CodeExplorerComponentViewModel(item, grouping)))
                        .OrderBy(item => item.Name)
                        .ToList();

            _customFolders = items.GroupBy(item => item.CustomFolder)
                .Select(grouping => new CodeExplorerCustomFolderViewModel(grouping.Key, grouping))
                .OrderBy(item => item.Name)
                .ToList();

            var isProtected = _declaration.Project.Protection == vbext_ProjectProtection.vbext_pp_locked;
            _icon = GetImageSource(resx.folder_horizontal);
        }

        public IEnumerable<CodeExplorerComponentViewModel> Components { get { return _components; } }
        public IEnumerable<CodeExplorerCustomFolderViewModel> CustomFolders { get { return _customFolders; } }

        public string Name { get { return _declaration.CustomFolder; } }
        public BitmapImage Icon { get { return _icon; } }
    }

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
        }

        public string Name { get { return _name; } }

        public IEnumerable<CodeExplorerComponentViewModel> Components { get { return _components; } }
    }
}