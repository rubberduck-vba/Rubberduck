using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Media.Imaging;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.Annotations;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;
using resx = Rubberduck.Resources.CodeExplorer.CodeExplorerUI;

namespace Rubberduck.Navigation.CodeExplorer
{
    public class CodeExplorerComponentViewModel : CodeExplorerItemViewModel, ICodeExplorerDeclarationViewModel
    {
        public Declaration Declaration { get; }

        public override CodeExplorerItemViewModel Parent { get; }

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

        private readonly IProjectsProvider _projectsProvider;

        public CodeExplorerComponentViewModel(CodeExplorerItemViewModel parent, Declaration declaration, IEnumerable<Declaration> declarations, IProjectsProvider projectsProvider)
        {
            Parent = parent;
            Declaration = declaration;
            _projectsProvider = projectsProvider;
            
            _icon = Icons.ContainsKey(DeclarationType) 
                ? Icons[DeclarationType]
                : GetImageSource(resx.status_offline);

            Items = declarations.GroupBy(item => item.Scope).SelectMany(grouping =>
                            grouping.Where(item => item.ParentDeclaration != null
                                                && item.ParentScope == declaration.Scope
                                                && MemberTypes.Contains(item.DeclarationType))
                                .OrderBy(item => item.QualifiedSelection.Selection.StartLine)
                                .Select(item => new CodeExplorerMemberViewModel(this, item, grouping)))
                                .ToList<CodeExplorerItemViewModel>();

            _name = Declaration.IdentifierName;
            var qualifiedModuleName = declaration.QualifiedName.QualifiedModuleName;

            try
            {
                switch (qualifiedModuleName.ComponentType)
                {
                    case ComponentType.Document:
                        var component = _projectsProvider.Component(qualifiedModuleName);
                        string parenthesizedName;
                        using (var properties = component.Properties)
                        {
                            parenthesizedName = properties["Name"].Value.ToString() ?? string.Empty;
                        }

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
                        break;

                    case ComponentType.ResFile:
                        var fileName = Declaration.IdentifierName.Split('\\').Last();
                        _name = $"{Resources.RubberduckUI.CodeExplorer_ResourceFileText} ({fileName})";
                        break;

                    case ComponentType.RelatedDocument:
                        _name = $"({Declaration.IdentifierName.Split('\\').Last()})";
                        break;

                    default:
                        _name = Declaration.IdentifierName;
                        break;
                }
            }
            catch
            {
                // gotcha! (this means that the property either doesn't exist or we weren't able to get it for some reason)
            }
        }

        private bool ContainsBuiltinDocumentPropertiesProperty()
        {
            using (var properties = _projectsProvider.Component(Declaration.QualifiedName.QualifiedModuleName).Properties)
            {
                return properties.Any(item => item.Name == "BuiltinDocumentProperties");
            }
        }

        private bool _isErrorState;
        public bool IsErrorState
        {
            get => _isErrorState;
            set
            {
                _isErrorState = value;
                _icon = GetImageSource(resx.cross_circle);


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
                return Declaration.DeclarationType == DeclarationType.ProceduralModule
                       && Declaration.Annotations.Any(annotation => annotation.AnnotationType == AnnotationType.TestModule);
            }
        }

        private readonly string _name;
        public override string Name => _name;
  
        public override string NameWithSignature => _name;

        public override QualifiedSelection? QualifiedSelection => Declaration.QualifiedSelection;

        private ComponentType ComponentType => Declaration.QualifiedName.QualifiedModuleName.ComponentType;

        private static readonly IDictionary<ComponentType, DeclarationType> DeclarationTypes = new Dictionary<ComponentType, DeclarationType>
        {
            { ComponentType.ClassModule, DeclarationType.ClassModule },
            { ComponentType.StandardModule, DeclarationType.ProceduralModule },
            { ComponentType.Document, DeclarationType.Document },
            { ComponentType.UserForm, DeclarationType.UserForm },
            { ComponentType.VBForm, DeclarationType.VbForm },
            { ComponentType.MDIForm, DeclarationType.MdiForm},
            { ComponentType.UserControl, DeclarationType.UserControl},
            { ComponentType.ResFile, DeclarationType.ResFile},
            { ComponentType.RelatedDocument, DeclarationType.RelatedDocument},
            { ComponentType.PropPage, DeclarationType.PropPage},
            { ComponentType.ActiveXDesigner, DeclarationType.ActiveXDesigner}
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
            { DeclarationType.Document, GetImageSource(resx.document_office) },
            { DeclarationType.VbForm, GetImageSource(resx.ProjectForm)},
            { DeclarationType.MdiForm, GetImageSource(resx.MdiForm)},
            { DeclarationType.UserControl, GetImageSource(resx.ui_scroll_pane_form)},
            { DeclarationType.DocObject, GetImageSource(resx.document_globe)},
            { DeclarationType.PropPage, GetImageSource(resx.ui_tab_content)},
            { DeclarationType.ActiveXDesigner, GetImageSource(resx.pencil_ruler)},
            { DeclarationType.ResFile, GetImageSource(resx.document_block)},
            { DeclarationType.RelatedDocument, GetImageSource(resx.document_import)}
        };

        private BitmapImage _icon;
        public override BitmapImage CollapsedIcon => _icon;
        public override BitmapImage ExpandedIcon => _icon;
    }
}
