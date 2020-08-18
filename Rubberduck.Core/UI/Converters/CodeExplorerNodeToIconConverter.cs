using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Data;
using System.Windows.Media;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Annotations.Concrete;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources.CodeExplorer;

namespace Rubberduck.UI.Converters
{
    public class CodeExplorerNodeToIconConverter : ImageSourceConverter, IMultiValueConverter
    {
        private static readonly ImageSource ExceptionIcon = ToImageSource(CodeExplorerUI.exclamation);
        private static readonly ImageSource ProjectIcon = ToImageSource(CodeExplorerUI.ObjectLibrary);
        private static readonly ImageSource InterfaceIcon = ToImageSource(CodeExplorerUI.ObjectInterface);
        private static readonly ImageSource PredeclaredIcon = ToImageSource(CodeExplorerUI.ObjectClassPredeclared);
        private static readonly ImageSource NullIcon = ToImageSource(CodeExplorerUI.status_offline);
        private static readonly ImageSource TestMethodIcon = ToImageSource(CodeExplorerUI.ObjectTestMethod);

        private static readonly ImageSource OpenFolderIcon = ToImageSource(CodeExplorerUI.FolderOpen);
        private static readonly ImageSource ClosedFolderIcon = ToImageSource(CodeExplorerUI.FolderClosed);

        private static readonly ImageSource ReferenceFolderIcon = ToImageSource(CodeExplorerUI.ObjectAssembly);
        private static readonly ImageSource ReferenceIcon = ToImageSource(CodeExplorerUI.Reference);
        private static readonly ImageSource LockedReferenceIcon = ToImageSource(CodeExplorerUI.LockedReference);
        private static readonly ImageSource BrokenReferenceIcon = ToImageSource(CodeExplorerUI.BrokenReference);

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
            { DeclarationType.Variable, ToImageSource(CodeExplorerUI.ObjectField)}
        };

        public override object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if ((value as ICodeExplorerNode)?.Declaration is null)
            {
                return NullIcon;
            }

            switch (value)
            {
                case CodeExplorerProjectViewModel _:
                    return ProjectIcon;
                case CodeExplorerReferenceFolderViewModel _:
                    return ReferenceFolderIcon;
                case CodeExplorerReferenceViewModel reference:
                    return reference.Reference is null || reference.Reference.IsBroken
                        ? BrokenReferenceIcon
                        : reference.Reference.IsBuiltIn
                            ? LockedReferenceIcon
                            : ReferenceIcon;
                case CodeExplorerCustomFolderViewModel folder:
                    return folder.IsExpanded ? OpenFolderIcon : ClosedFolderIcon;
                case CodeExplorerComponentViewModel component:
                    if (component.IsPredeclared)
                    {
                        return PredeclaredIcon;
                    }

                    if (component.Declaration is null)
                    {
                        return ExceptionIcon;
                    }

                    if (component.Declaration is ClassModuleDeclaration classModule && 
                        (classModule.IsInterface || classModule.Annotations.Any(annotation => annotation.Annotation is InterfaceAnnotation)))
                    {
                        return InterfaceIcon;
                    }

                    var isUserForm = component.Declaration.QualifiedModuleName.ComponentType == VBEditor.SafeComWrappers.ComponentType.UserForm; 
                    return DeclarationIcons.ContainsKey(component.Declaration.DeclarationType)
                        ? DeclarationIcons[isUserForm ? DeclarationType.UserForm : component.Declaration.DeclarationType]
                        : ExceptionIcon;
                default:
                    return value is ICodeExplorerNode node &&
                           node.Declaration != null &&
                           DeclarationIcons.ContainsKey(node.Declaration.DeclarationType)
                        ? node.Declaration.Annotations.Any(a => a.Annotation is TestMethodAnnotation)
                            ? TestMethodIcon
                            : DeclarationIcons[node.Declaration.DeclarationType]
                        : ExceptionIcon;
            }
        }

        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values.Length == 0)
            {
                return null;
            };

            return Convert(values[0], targetType, parameter, culture);
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
