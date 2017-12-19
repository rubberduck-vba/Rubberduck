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
        private static readonly DeclarationType[] ComponentTypes =
        {
            DeclarationType.ClassModule, 
            DeclarationType.Document, 
            DeclarationType.ProceduralModule, 
            DeclarationType.UserForm, 
        };

        public CodeExplorerCustomFolderViewModel(CodeExplorerItemViewModel parent, string name, string fullPath)
        {
            _parent = parent;
            FullPath = fullPath;
            Name = name.Replace("\"", string.Empty);
            FolderAttribute = string.Format("@Folder(\"{0}\")", fullPath.Replace("\"", string.Empty));

            CollapsedIcon = GetImageSource(resx.folder_horizontal);
            ExpandedIcon = GetImageSource(resx.folder_horizontal_open);
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

                    AddChild(new CodeExplorerComponentViewModel(this, parent, members));
                }
                catch (InvalidOperationException exception)
                {
                    Console.WriteLine(exception);
                }
            }
        }

        public string FolderAttribute { get; }

        public string FullPath { get; }

        public override string Name { get; }

        public override string NameWithSignature => Name; // Is this actually doing anything? Should this member be replaced with 'Name'?

        public override QualifiedSelection? QualifiedSelection => null;

        public override BitmapImage CollapsedIcon { get; }

        public override BitmapImage ExpandedIcon { get; }

        // I have to set the parent from a different location than
        // the node is created because of the folder helper
        internal void SetParent(CodeExplorerItemViewModel parent)
        {
            _parent = parent;
        }

        private CodeExplorerItemViewModel _parent;
        public override CodeExplorerItemViewModel Parent => _parent;
    }
}
