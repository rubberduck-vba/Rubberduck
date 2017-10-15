using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Media.Imaging;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using resx = Rubberduck.UI.CodeExplorer.CodeExplorer;
using Rubberduck.Parsing.Annotations;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Navigation.CodeExplorer
{
    public class CodeExplorerComponentViewModel : CodeExplorerItemViewModel, ICodeExplorerDeclarationViewModel
    {
        private readonly Declaration _declaration;
        public Declaration Declaration { get { return _declaration; } }

        private readonly CodeExplorerItemViewModel _parent;
        public override CodeExplorerItemViewModel Parent { get { return _parent; } }

        private static readonly DeclarationType[] MemberTypes =
        {
            DeclarationType.Constant, 
            DeclarationType.Enumeration, 
            DeclarationType.Event, 
            DeclarationType.Function, 
            DeclarationType.LibraryFunction, 
            DeclarationType.LibraryProcedure, 
            DeclarationType.Procedure,
            DeclarationType.PropertyGet, 
            DeclarationType.PropertyLet, 
            DeclarationType.PropertySet, 
            DeclarationType.UserDefinedType, 
            DeclarationType.Variable, 
        };

        public CodeExplorerComponentViewModel(CodeExplorerItemViewModel parent, Declaration declaration, IEnumerable<Declaration> declarations)
        {
            _parent = parent;
            _declaration = declaration;
            _icon = Icons[DeclarationType];
            Items = declarations.GroupBy(item => item.Scope).SelectMany(grouping =>
                            grouping.Where(item => item.ParentDeclaration != null
                                                && item.ParentScope == declaration.Scope
                                                && MemberTypes.Contains(item.DeclarationType))
                                .OrderBy(item => item.QualifiedSelection.Selection.StartLine)
                                .Select(item => new CodeExplorerMemberViewModel(this, item, grouping)))
                                .ToList<CodeExplorerItemViewModel>();

            _name = _declaration.IdentifierName;

            var component = declaration.QualifiedName.QualifiedModuleName.Component;
            if (component.Type == ComponentType.Document)
            {
                try
                {
                    var parenthesizedName = component.Properties["Name"].Value.ToString();

                    if (ContainsBuiltinDocumentPropertiesProperty())
                    {
                        CodeExplorerItemViewModel node = this;
                        while (node.Parent != null)
                        {
                            node = node.Parent;
                        }

                        ((CodeExplorerProjectViewModel) node).SetParenthesizedName(parenthesizedName);
                    }
                    else
                    {
                        _name += " (" + parenthesizedName + ")";
                    }
                }
                catch
                {
                    // gotcha! (this means that the property either doesn't exist or we weren't able to get it for some reason)
                }
            }
        }

        private bool ContainsBuiltinDocumentPropertiesProperty()
        {
            var properties = _declaration.QualifiedName.QualifiedModuleName.Component.Properties;
            {
                return properties.Any(item => item.Name == "BuiltinDocumentProperties");
            }
        }

        private bool _isErrorState;
        public bool IsErrorState
        {
            get { return _isErrorState; }
            set
            {
                _isErrorState = value;
                _icon = GetImageSource(resx.Error);


                foreach (var item in Items)
                {
                    ((CodeExplorerMemberViewModel) item).ParentComponentHasError();
                }

                OnPropertyChanged();
                OnPropertyChanged("CollapsedIcon");
                OnPropertyChanged("ExpandedIcon");
            }
        }

        public bool IsTestModule
        {
            get
            {
                return _declaration.DeclarationType == DeclarationType.ProceduralModule
                       && _declaration.Annotations.Any(annotation => annotation.AnnotationType == AnnotationType.TestModule);
            }
        }

        private readonly string _name;
        public override string Name { get { return _name; } }
        public override string NameWithSignature { get { return _name; } }

        public override QualifiedSelection? QualifiedSelection { get { return _declaration.QualifiedSelection; } }

        private ComponentType ComponentType { get { return _declaration.QualifiedName.QualifiedModuleName.ComponentType; } }

        private static readonly IDictionary<ComponentType, DeclarationType> DeclarationTypes = new Dictionary<ComponentType, DeclarationType>
        {
            { ComponentType.ClassModule, DeclarationType.ClassModule },
            { ComponentType.StandardModule, DeclarationType.ProceduralModule },
            { ComponentType.Document, DeclarationType.Document },
            { ComponentType.UserForm, DeclarationType.UserForm }
        };

        private DeclarationType DeclarationType
        {
            get
            {
                var result = DeclarationType.ClassModule;
                try
                {
                    DeclarationTypes.TryGetValue(ComponentType, out result);
                }
                catch (COMException exception)
                {
                    Console.WriteLine(exception);
                }
                return result;
            }
        }

        private static readonly IDictionary<DeclarationType,BitmapImage> Icons = new Dictionary<DeclarationType, BitmapImage>
        {
            { DeclarationType.ClassModule, GetImageSource(resx.ObjectClass) },
            { DeclarationType.ProceduralModule, GetImageSource(resx.ObjectModule) },
            { DeclarationType.UserForm, GetImageSource(resx.ProjectForm) },
            { DeclarationType.Document, GetImageSource(resx.document_office) }
        };

        private BitmapImage _icon;
        public override BitmapImage CollapsedIcon { get { return _icon; } }
        public override BitmapImage ExpandedIcon { get { return _icon; } }
    }
}
