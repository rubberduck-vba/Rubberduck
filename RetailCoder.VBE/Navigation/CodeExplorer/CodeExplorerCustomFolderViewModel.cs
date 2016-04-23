using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Media.Imaging;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using resx = Rubberduck.Properties.Resources;

namespace Rubberduck.Navigation.CodeExplorer
{
    public class CodeExplorerCustomFolderViewModel : CodeExplorerItemViewModel
    {
        private readonly string _name;
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

            _collapsedIcon = GetImageSource(resx.folder_horizontal);
            _expandedIcon = GetImageSource(resx.folder_horizontal_open);

            var items = declarations.ToList();

            var parents = items.GroupBy(item => item.ComponentName).OrderBy(item => item.Key).ToList();
            foreach (var component in parents)
            {
                try
                {
                    var moduleName = component.Key;
                    var parent = items.Single(item =>
                        ComponentTypes.Contains(item.DeclarationType) && item.ComponentName == moduleName);
                    var members = items.Where(item =>
                        !ComponentTypes.Contains(item.DeclarationType) && item.ComponentName == moduleName);

                    AddChild(new CodeExplorerComponentViewModel(parent, members));
                }
                catch (InvalidOperationException exception)
                {
                    Console.WriteLine(exception);
                }
            }
        }

        public override string Name { get { return _name; } }

        public override QualifiedSelection? QualifiedSelection { get { return null; } }

        private readonly BitmapImage _collapsedIcon;
        public override BitmapImage CollapsedIcon { get { return _collapsedIcon; } }

        private readonly BitmapImage _expandedIcon;
        public override BitmapImage ExpandedIcon { get { return _expandedIcon; } }
    }
}