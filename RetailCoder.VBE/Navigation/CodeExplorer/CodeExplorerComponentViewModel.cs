using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Media.Imaging;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using resx = Rubberduck.UI.CodeExplorer.CodeExplorer;
using Rubberduck.Parsing.Annotations;

namespace Rubberduck.Navigation.CodeExplorer
{
    public class CodeExplorerComponentViewModel : CodeExplorerItemViewModel
    {
        private readonly Declaration _declaration;
        public Declaration Declaration { get { return _declaration; } }

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

        public CodeExplorerComponentViewModel(Declaration declaration, IEnumerable<Declaration> declarations)
        {
            _declaration = declaration;
            _icon = Icons[DeclarationType];
            Items = declarations.GroupBy(item => item.Scope).SelectMany(grouping =>
                            grouping.Where(item => item.ParentDeclaration != null
                                                && item.ParentScope == declaration.Scope
                                                && MemberTypes.Contains(item.DeclarationType))
                                .OrderBy(item => item.QualifiedSelection.Selection.StartLine)
                                .Select(item => new CodeExplorerMemberViewModel(item, grouping)))
                                .ToList<CodeExplorerItemViewModel>();
            
        }

        private bool _isErrorState;
        public bool IsErrorState { get { return _isErrorState; } set { _isErrorState = value; OnPropertyChanged(); } }

        public bool IsTestModule
        {
            get
            {
                return _declaration.DeclarationType == DeclarationType.ProceduralModule
                       && _declaration.Annotations.Any(annotation => annotation.AnnotationType == AnnotationType.TestModule);
            }
        }

        public override string Name { get { return _declaration.IdentifierName; } }
        public override string NameWithSignature { get { return Name; } }

        public override QualifiedSelection? QualifiedSelection { get { return _declaration.QualifiedSelection; } }

        private vbext_ComponentType ComponentType { get { return _declaration.QualifiedName.QualifiedModuleName.Component.Type; } }

        private static readonly IDictionary<vbext_ComponentType, DeclarationType> DeclarationTypes = new Dictionary<vbext_ComponentType, DeclarationType>
        {
            { vbext_ComponentType.vbext_ct_ClassModule, DeclarationType.ClassModule },
            { vbext_ComponentType.vbext_ct_StdModule, DeclarationType.ProceduralModule },
            { vbext_ComponentType.vbext_ct_Document, DeclarationType.Document },
            { vbext_ComponentType.vbext_ct_MSForm, DeclarationType.UserForm }
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
            { DeclarationType.ClassModule, GetImageSource(resx.VSObject_Class) },
            { DeclarationType.ProceduralModule, GetImageSource(resx.VSObject_Module) },
            { DeclarationType.UserForm, GetImageSource(resx.VSProject_form) },
            { DeclarationType.Document, GetImageSource(resx.document_office) }
        };

        private readonly BitmapImage _icon;
        public override BitmapImage CollapsedIcon { get { return _icon; } }
        public override BitmapImage ExpandedIcon { get { return _icon; } }
    }
}