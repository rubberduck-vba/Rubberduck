using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Media.Imaging;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using resx = Rubberduck.UI.CodeExplorer.CodeExplorer;

namespace Rubberduck.Navigation.CodeExplorer
{
    public class CodeExplorerMemberViewModel : ViewModelBase
    {
        private readonly Declaration _declaration;

        private static readonly DeclarationType[] SubMemberTypes =
        {
            DeclarationType.EnumerationMember, 
            DeclarationType.UserDefinedTypeMember            
        };

        private static readonly IDictionary<Tuple<DeclarationType,Accessibility>,BitmapImage> Mappings =
            new Dictionary<Tuple<DeclarationType, Accessibility>, BitmapImage>
            {
                { Tuple.Create(DeclarationType.Constant, Accessibility.Private), GetImageSource(resx.VSObject_Constant_Private)},
                { Tuple.Create(DeclarationType.Constant, Accessibility.Public), GetImageSource(resx.VSObject_Field)},
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

        public CodeExplorerMemberViewModel(Declaration declaration, IEnumerable<Declaration> declarations)
        {
            _declaration = declaration;
            if (declarations != null)
            {
                _members = declarations.Where(item => SubMemberTypes.Contains(item.DeclarationType) && item.ParentDeclaration.Equals(declaration))
                                       .Select(item => new CodeExplorerMemberViewModel(item, null))
                                       .OrderBy(item => item.Name);
            }

            var modifier = declaration.Accessibility == Accessibility.Global || declaration.Accessibility == Accessibility.Implicit
                ? Accessibility.Public
                : declaration.Accessibility;
            var key = Tuple.Create(declaration.DeclarationType, modifier);
            _icon = Mappings[key];
        }

        public string Name { get { return _declaration.IdentifierName; } }

        private readonly BitmapImage _icon;
        public BitmapImage Icon { get { return _icon; } }

        private readonly IEnumerable<CodeExplorerMemberViewModel> _members;
        public IEnumerable<CodeExplorerMemberViewModel> Members { get { return _members; } }
    }
}