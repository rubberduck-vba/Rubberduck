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
            _components = declarations.GroupBy(item => item.ComponentName)
                .SelectMany(grouping =>
                    grouping.Where(item => ComponentTypes.Contains(item.DeclarationType))
                        .Select(item => new CodeExplorerComponentViewModel(item, grouping)))
                        .OrderBy(item => item.Name)
                        .ToList();

            var isProtected = _declaration.Project.Protection == vbext_ProjectProtection.vbext_pp_locked;
            _icon = GetImageSource(resx.folder_horizontal);
        }

        public IEnumerable<CodeExplorerComponentViewModel> Components { get { return _components; } }

        public string Name { get { return _declaration.IdentifierName; } }
        public BitmapImage Icon { get { return _icon; } }
    }
}