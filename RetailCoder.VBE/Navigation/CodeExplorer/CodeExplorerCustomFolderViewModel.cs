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
        private readonly string _fullPath;
        private readonly string _name;
        private readonly string _folderAttribute;
        private static readonly DeclarationType[] ComponentTypes =
        {
            DeclarationType.ClassModule, 
            DeclarationType.Document, 
            DeclarationType.ProceduralModule, 
            DeclarationType.UserForm, 
        };

        public CodeExplorerCustomFolderViewModel(string name, string fullPath)
        {
            _fullPath = fullPath;
            _name = name.Replace("\"", string.Empty);
            _folderAttribute = string.Format("@Folder(\"{0}\")", fullPath.Replace("\"", string.Empty));

            _collapsedIcon = GetImageSource(resx.folder_horizontal);
            _expandedIcon = GetImageSource(resx.folder_horizontal_open);
        }

        public void AddNodes(List<Declaration> declarations)
        {
            var parents = declarations.GroupBy(item => item.ComponentName).OrderBy(item => item.Key).ToList();
            foreach (var component in parents)
            {
                try
                {
                    var moduleName = component.Key;
                    var parent = declarations.Single(item =>
                        ComponentTypes.Contains(item.DeclarationType) && item.ComponentName == moduleName);
                    var members = declarations.Where(item =>
                        !ComponentTypes.Contains(item.DeclarationType) && item.ComponentName == moduleName);

                    AddChild(new CodeExplorerComponentViewModel(parent, members));
                }
                catch (InvalidOperationException exception)
                {
                    Console.WriteLine(exception);
                }
            }
        }

        public string FolderAttribute { get { return _folderAttribute; } }

        public string FullPath { get { return _fullPath; } }

        public override string Name { get { return _name; } }
        public override string NameWithSignature { get { return Name; } }

        public override QualifiedSelection? QualifiedSelection { get { return null; } }

        private readonly BitmapImage _collapsedIcon;
        public override BitmapImage CollapsedIcon { get { return _collapsedIcon; } }

        private readonly BitmapImage _expandedIcon;
        public override BitmapImage ExpandedIcon { get { return _expandedIcon; } }
    }
}