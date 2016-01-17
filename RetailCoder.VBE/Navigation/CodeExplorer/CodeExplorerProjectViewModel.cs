using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Media.Imaging;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using resx = Rubberduck.UI.CodeExplorer.CodeExplorer;

namespace Rubberduck.Navigation.CodeExplorer
{
    public class CodeExplorerProjectViewModel : ViewModelBase
    {
        private readonly Declaration _declaration;
        private readonly IEnumerable<CodeExplorerComponentViewModel> _components;
        private readonly Lazy<IEnumerable<CodeExplorerCustomFolderViewModel>> _customFolders; 

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

            _customFolders = new Lazy<IEnumerable<CodeExplorerCustomFolderViewModel>>(() => 
                items.GroupBy(item => item.CustomFolder)
                     .Select(grouping => new CodeExplorerCustomFolderViewModel(grouping.Key, grouping))
                     .OrderBy(item => item.Name)
                     .ToList());

            _icon = _declaration.Project.Protection == vbext_ProjectProtection.vbext_pp_locked
                        ? GetImageSource(resx.lock__exclamation)
                        : GetImageSource(resx.VSObject_Library);
        }

        private readonly BitmapImage _icon;
        public BitmapImage Icon { get { return _icon; } }

        public IEnumerable<CodeExplorerComponentViewModel> Components { get { return _components; } }
        public IEnumerable<CodeExplorerCustomFolderViewModel> CustomFolders { get { return _customFolders.Value; } }

        public string Name { get { return _declaration.CustomFolder; } }
    }
}