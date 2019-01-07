using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Media.Imaging;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.Annotations;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.Resources.CodeExplorer;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

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
        private readonly IVBE _vbe;

        public CodeExplorerComponentViewModel(CodeExplorerItemViewModel parent, Declaration declaration, IEnumerable<Declaration> declarations, IProjectsProvider projectsProvider, IVBE vbe)
        {
            Parent = parent;
            Declaration = declaration;
            _projectsProvider = projectsProvider;
            _vbe = vbe;

            _icon = Icons.ContainsKey(DeclarationType) 
                ? Icons[DeclarationType]
                : GetImageSource(CodeExplorerUI.status_offline);

            Items = declarations.GroupBy(item => item.Scope).SelectMany(grouping =>
                            grouping.Where(item => item.ParentDeclaration != null
                                                && item.ParentScope == declaration.Scope
                                                && MemberTypes.Contains(item.DeclarationType))
                                .OrderBy(item => item.QualifiedSelection.Selection.StartLine)
                                .Select(item => new CodeExplorerMemberViewModel(this, item, grouping)))
                                .ToList<CodeExplorerItemViewModel>();

            _name = DeclarationType == DeclarationType.ResFile && string.IsNullOrEmpty(Declaration.IdentifierName) 
                ? CodeExplorerUI.CodeExplorer_ResourceFileText
                : Declaration.IdentifierName;

            var qualifiedModuleName = declaration.QualifiedName.QualifiedModuleName;
            try
            {
                switch (qualifiedModuleName.ComponentType)
                {
                    case ComponentType.Document:
                        var parenthesizedName = string.Empty;
                        var state = DocumentState.Inaccessible;
                        using (var app = _vbe.HostApplication())
                        {
                            if (app != null)
                            {
                                var document = app.GetDocument(qualifiedModuleName);
                                parenthesizedName = document?.DocumentName ?? string.Empty;
                                state = document?.State ?? DocumentState.Inaccessible;
                            }
                        }
                        
                        if (state == DocumentState.DesignView && ContainsBuiltinDocumentPropertiesProperty())
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
                            if (!string.IsNullOrWhiteSpace(parenthesizedName))
                            {
                                _name += " (" + parenthesizedName + ")";
                            }
                        }
                        break;

                    case ComponentType.ResFile:
                        var fileName = Declaration.IdentifierName.Split('\\').Last();
                        _name = $"{CodeExplorerUI.CodeExplorer_ResourceFileText} ({fileName})";
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
            var component = _projectsProvider.Component(Declaration.QualifiedName.QualifiedModuleName);
            using (var properties = component.Properties)
            {
                foreach (var property in properties)
                using(property)
                {
                    if (property.Name == "BuiltinDocumentProperties")
                    {
                        return true;
                    }
                }

                return false;
            }
        }

        private bool _isErrorState;
        public bool IsErrorState
        {
            get => _isErrorState;
            set
            {
                _isErrorState = value;
                _icon = GetImageSource(CodeExplorerUI.cross_circle);


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
            { ComponentType.DocObject, DeclarationType.DocObject},
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
            { DeclarationType.ClassModule, GetImageSource(CodeExplorerUI.ObjectClass) },
            { DeclarationType.ProceduralModule, GetImageSource(CodeExplorerUI.ObjectModule) },
            { DeclarationType.UserForm, GetImageSource(CodeExplorerUI.ProjectForm) },
            { DeclarationType.Document, GetImageSource(CodeExplorerUI.document_office) },
            { DeclarationType.VbForm, GetImageSource(CodeExplorerUI.ProjectForm)},
            { DeclarationType.MdiForm, GetImageSource(CodeExplorerUI.MdiForm)},
            { DeclarationType.UserControl, GetImageSource(CodeExplorerUI.ui_scroll_pane_form)},
            { DeclarationType.DocObject, GetImageSource(CodeExplorerUI.document_globe)},
            { DeclarationType.PropPage, GetImageSource(CodeExplorerUI.ui_tab_content)},
            { DeclarationType.ActiveXDesigner, GetImageSource(CodeExplorerUI.pencil_ruler)},
            { DeclarationType.ResFile, GetImageSource(CodeExplorerUI.document_block)},
            { DeclarationType.RelatedDocument, GetImageSource(CodeExplorerUI.document_import)}
        };

        private BitmapImage _icon;
        public override BitmapImage CollapsedIcon => _icon;
        public override BitmapImage ExpandedIcon => _icon;
    }
}
