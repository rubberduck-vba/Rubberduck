using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Media.Imaging;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using resx = Rubberduck.UI.CodeExplorer.CodeExplorer;

namespace Rubberduck.Navigation.CodeExplorer
{
    public class CodeExplorerMemberViewModel : CodeExplorerItemViewModel, ICodeExplorerDeclarationViewModel
    {
        private readonly Declaration _declaration;
        public Declaration Declaration { get { return _declaration; } }

        private static readonly DeclarationType[] SubMemberTypes =
        {
            DeclarationType.EnumerationMember, 
            DeclarationType.UserDefinedTypeMember            
        };

        private static readonly IDictionary<Tuple<DeclarationType,Accessibility>,BitmapImage> Mappings =
            new Dictionary<Tuple<DeclarationType, Accessibility>, BitmapImage>
            {
                { Tuple.Create(DeclarationType.Constant, Accessibility.Private), GetImageSource(resx.VSObject_Constant_Private)},
                { Tuple.Create(DeclarationType.Constant, Accessibility.Public), GetImageSource(resx.VSObject_Constant)},
                { Tuple.Create(DeclarationType.Enumeration, Accessibility.Public), GetImageSource(resx.VSObject_Enum)},
                { Tuple.Create(DeclarationType.Enumeration, Accessibility.Private ), GetImageSource(resx.VSObject_EnumPrivate)},
                { Tuple.Create(DeclarationType.EnumerationMember, Accessibility.Public), GetImageSource(resx.VSObject_EnumItem)},
                { Tuple.Create(DeclarationType.Event, Accessibility.Public), GetImageSource(resx.VSObject_Event)},
                { Tuple.Create(DeclarationType.Event, Accessibility.Private ), GetImageSource(resx.VSObject_Event_Private)},
                { Tuple.Create(DeclarationType.Function, Accessibility.Public), GetImageSource(resx.VSObject_Method)},
                { Tuple.Create(DeclarationType.Function, Accessibility.Friend ), GetImageSource(resx.VSObject_Method_Friend)},
                { Tuple.Create(DeclarationType.Function, Accessibility.Private ), GetImageSource(resx.VSObject_Method_Private)},
                { Tuple.Create(DeclarationType.LibraryFunction, Accessibility.Public), GetImageSource(resx.VSObject_Method_Shortcut)},
                { Tuple.Create(DeclarationType.LibraryProcedure, Accessibility.Public), GetImageSource(resx.VSObject_Method_Shortcut)},
                { Tuple.Create(DeclarationType.LibraryFunction, Accessibility.Private), GetImageSource(resx.VSObject_Method_Shortcut)},
                { Tuple.Create(DeclarationType.LibraryProcedure, Accessibility.Private), GetImageSource(resx.VSObject_Method_Shortcut)},
                { Tuple.Create(DeclarationType.LibraryFunction, Accessibility.Friend), GetImageSource(resx.VSObject_Method_Shortcut)},
                { Tuple.Create(DeclarationType.LibraryProcedure, Accessibility.Friend), GetImageSource(resx.VSObject_Method_Shortcut)},
                { Tuple.Create(DeclarationType.Procedure, Accessibility.Public), GetImageSource(resx.VSObject_Method)},
                { Tuple.Create(DeclarationType.Procedure, Accessibility.Friend ), GetImageSource(resx.VSObject_Method_Friend)},
                { Tuple.Create(DeclarationType.Procedure, Accessibility.Private ), GetImageSource(resx.VSObject_Method_Private)},
                { Tuple.Create(DeclarationType.PropertyGet, Accessibility.Public), GetImageSource(resx.VSObject_Properties)},
                { Tuple.Create(DeclarationType.PropertyGet, Accessibility.Friend ), GetImageSource(resx.VSObject_Properties_Friend)},
                { Tuple.Create(DeclarationType.PropertyGet, Accessibility.Private ), GetImageSource(resx.VSObject_Properties_Private)},
                { Tuple.Create(DeclarationType.PropertyLet, Accessibility.Public), GetImageSource(resx.VSObject_Properties)},
                { Tuple.Create(DeclarationType.PropertyLet, Accessibility.Friend ), GetImageSource(resx.VSObject_Properties_Friend)},
                { Tuple.Create(DeclarationType.PropertyLet, Accessibility.Private ), GetImageSource(resx.VSObject_Properties_Private)},
                { Tuple.Create(DeclarationType.PropertySet, Accessibility.Public), GetImageSource(resx.VSObject_Properties)},
                { Tuple.Create(DeclarationType.PropertySet, Accessibility.Friend ), GetImageSource(resx.VSObject_Properties_Friend)},
                { Tuple.Create(DeclarationType.PropertySet, Accessibility.Private ), GetImageSource(resx.VSObject_Properties_Private)},
                { Tuple.Create(DeclarationType.UserDefinedType, Accessibility.Public), GetImageSource(resx.VSObject_ValueType)},
                { Tuple.Create(DeclarationType.UserDefinedType, Accessibility.Private ), GetImageSource(resx.VSObject_ValueTypePrivate)},
                { Tuple.Create(DeclarationType.UserDefinedTypeMember, Accessibility.Public), GetImageSource(resx.VSObject_Field)},
                { Tuple.Create(DeclarationType.Variable, Accessibility.Private), GetImageSource(resx.VSObject_Field_Private)},
                { Tuple.Create(DeclarationType.Variable, Accessibility.Public ), GetImageSource(resx.VSObject_Field)},
            };

        public CodeExplorerMemberViewModel(CodeExplorerItemViewModel parent, Declaration declaration, IEnumerable<Declaration> declarations)
        {
            _parent = parent;

            _declaration = declaration;
            if (declarations != null)
            {
                Items = declarations.Where(item => SubMemberTypes.Contains(item.DeclarationType) && item.ParentDeclaration.Equals(declaration))
                                    .OrderBy(item => item.Selection.StartLine)
                                    .Select(item => new CodeExplorerMemberViewModel(this, item, null))
                                    .ToList<CodeExplorerItemViewModel>();
            }

            var modifier = declaration.Accessibility == Accessibility.Global || declaration.Accessibility == Accessibility.Implicit
                ? Accessibility.Public
                : declaration.Accessibility;
            var key = Tuple.Create(declaration.DeclarationType, modifier);

            _name = DetermineMemberName(declaration);
            _icon = Mappings[key];
        }

        private readonly string _name;
        public override string Name { get { return _name; } }

        private readonly CodeExplorerItemViewModel _parent;
        public override CodeExplorerItemViewModel Parent { get { return _parent; } }

        private string _signature = null;
        public override string NameWithSignature
        {
            get
            {
                if (_signature != null)
                {
                    return _signature;
                }

                var context =
                    _declaration.Context.children.FirstOrDefault(d => d is VBAParser.ArgListContext) as VBAParser.ArgListContext;

                if (context == null)
                {
                    _signature = Name;
                }
                else if (_declaration.DeclarationType == DeclarationType.PropertyGet 
                      || _declaration.DeclarationType == DeclarationType.PropertyLet 
                      || _declaration.DeclarationType == DeclarationType.PropertySet)
                {
                    // 6 being the three-letter "get/let/set" + parens + space
                    _signature = Name.Insert(Name.Length - 6, context.GetText()); 
                }
                else
                {
                    _signature = Name + context.GetText();
                }
                return _signature;
            }
        }

        public override QualifiedSelection? QualifiedSelection { get { return _declaration.QualifiedSelection; } }

        private static string DetermineMemberName(Declaration declaration)
        {
            var type = declaration.DeclarationType;
            switch (type)
            {
                case DeclarationType.PropertyGet:
                    return declaration.IdentifierName + " (Get)";
                case DeclarationType.PropertyLet:
                    return declaration.IdentifierName + " (Let)";
                case DeclarationType.PropertySet:
                    return declaration.IdentifierName + " (Set)";
                case DeclarationType.Variable:
                    if (declaration.IsArray)
                    {
                        return declaration.IdentifierName + "()";
                    }
                    return declaration.IdentifierName;
                case DeclarationType.Constant:
                    var valuedDeclaration = (ConstantDeclaration)declaration;
                    return valuedDeclaration.IdentifierName + " = " + valuedDeclaration.Expression;

                default:
                    return declaration.IdentifierName;
            }
        }

        public void ParentComponentHasError()
        {
            _icon = GetImageSource(resx.Warning);
            OnPropertyChanged("CollapsedIcon");
            OnPropertyChanged("ExpandedIcon");
        }

        private BitmapImage _icon;
        public override BitmapImage CollapsedIcon { get { return _icon; } }
        public override BitmapImage ExpandedIcon { get { return _icon; } }
    }
}
