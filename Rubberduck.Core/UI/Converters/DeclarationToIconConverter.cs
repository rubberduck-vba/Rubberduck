using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Media;
using Rubberduck.Parsing.Annotations.Concrete;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources.CodeExplorer;

namespace Rubberduck.UI.Converters
{
    public class DeclarationToIconConverter : ImageSourceConverter
    {
        private static readonly ImageSource NullIcon = ToImageSource(CodeExplorerUI.status_offline);
        private static readonly ImageSource ExceptionIcon = ToImageSource(CodeExplorerUI.exclamation);
        private static readonly ImageSource InterfaceIcon = ToImageSource(CodeExplorerUI.ObjectInterface);
        private static readonly ImageSource PredeclaredIcon = ToImageSource(CodeExplorerUI.ObjectClassPredeclared);
        private static readonly ImageSource TestMethodIcon = ToImageSource(CodeExplorerUI.ObjectTestMethod);

        protected ImageSource NullIconSource => NullIcon;
        protected ImageSource ExceptionIconSource => ExceptionIcon;

        private static readonly IDictionary<DeclarationType, ImageSource> DeclarationIcons = new Dictionary<DeclarationType, ImageSource>
        {
            // Components
            { DeclarationType.ClassModule, ToImageSource(CodeExplorerUI.ObjectClass) },
            { DeclarationType.ProceduralModule, ToImageSource(CodeExplorerUI.ObjectModule) },
            { DeclarationType.UserForm, ToImageSource(CodeExplorerUI.ProjectForm) },
            { DeclarationType.Document, ToImageSource(CodeExplorerUI.document_office) },
            { DeclarationType.VbForm, ToImageSource(CodeExplorerUI.ProjectForm)},
            { DeclarationType.MdiForm, ToImageSource(CodeExplorerUI.MdiForm)},
            { DeclarationType.UserControl, ToImageSource(CodeExplorerUI.ui_scroll_pane_form)},
            { DeclarationType.DocObject, ToImageSource(CodeExplorerUI.document_globe)},
            { DeclarationType.PropPage, ToImageSource(CodeExplorerUI.ui_tab_content)},
            { DeclarationType.ActiveXDesigner, ToImageSource(CodeExplorerUI.pencil_ruler)},
            { DeclarationType.ResFile, ToImageSource(CodeExplorerUI.document_block)},
            { DeclarationType.RelatedDocument, ToImageSource(CodeExplorerUI.document_import)},          
            // Members
            { DeclarationType.Constant, ToImageSource(CodeExplorerUI.ObjectConstant)},
            { DeclarationType.Enumeration, ToImageSource(CodeExplorerUI.ObjectEnum)},
            { DeclarationType.EnumerationMember, ToImageSource(CodeExplorerUI.ObjectEnumItem)},
            { DeclarationType.Event, ToImageSource(CodeExplorerUI.ObjectEvent)},
            { DeclarationType.Function, ToImageSource(CodeExplorerUI.ObjectMethod)},
            { DeclarationType.LibraryFunction, ToImageSource(CodeExplorerUI.ObjectLibraryFunction)},
            { DeclarationType.LibraryProcedure, ToImageSource(CodeExplorerUI.ObjectLibraryFunction)},
            { DeclarationType.Procedure, ToImageSource(CodeExplorerUI.ObjectMethod)},
            { DeclarationType.PropertyGet, ToImageSource(CodeExplorerUI.ObjectPropertyGet)},
            { DeclarationType.PropertyLet, ToImageSource(CodeExplorerUI.ObjectPropertyLet)},
            { DeclarationType.PropertySet, ToImageSource(CodeExplorerUI.ObjectPropertySet)},
            { DeclarationType.UserDefinedType, ToImageSource(CodeExplorerUI.ObjectValueType)},
            { DeclarationType.UserDefinedTypeMember, ToImageSource(CodeExplorerUI.ObjectField)},
            { DeclarationType.Variable, ToImageSource(CodeExplorerUI.ObjectField)},
            { DeclarationType.Parameter, ToImageSource(CodeExplorerUI.ObjectField)},
        };

        public override object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null)
            {
                return NullIcon;
            }

            if (value is Declaration declaration)
            {
                if (declaration is ClassModuleDeclaration classModule)
                {
                    if (classModule.QualifiedModuleName.ComponentType ==
                        VBEditor.SafeComWrappers.ComponentType.Document)
                    {
                        return DeclarationIcons[DeclarationType.Document];
                    }
                    
                    if (classModule.QualifiedModuleName.ComponentType == VBEditor.SafeComWrappers.ComponentType.UserForm)
                    {
                        // a form has a predeclared ID, but we want it to have a UserForm icon:
                        return DeclarationIcons[DeclarationType.UserForm];
                    }

                    if (classModule.IsInterface || classModule.Annotations.Any(annotation => annotation.Annotation is InterfaceAnnotation))
                    {
                        return InterfaceIcon;
                    }

                    if (classModule.HasPredeclaredId)
                    {
                        return PredeclaredIcon;
                    }
                    
                    return DeclarationIcons.ContainsKey(classModule.DeclarationType)
                        ? DeclarationIcons[classModule.DeclarationType]
                        : NullIcon;
                }

                if (DeclarationIcons.ContainsKey(declaration.DeclarationType))
                {
                    if (declaration.Annotations.Any(a => a.Annotation is TestMethodAnnotation))
                    {
                        return TestMethodIcon;
                    }

                    return DeclarationIcons[declaration.DeclarationType];
                }
                return NullIcon;
            }

            return null;
        }
    }
}