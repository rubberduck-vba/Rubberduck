using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
                { Tuple.Create(DeclarationType.Constant, Accessibility.Private), GetImageSource(resx.ObjectConstantPrivate)},
                { Tuple.Create(DeclarationType.Constant, Accessibility.Public), GetImageSource(resx.ObjectConstant)},
                { Tuple.Create(DeclarationType.Enumeration, Accessibility.Public), GetImageSource(resx.ObjectEnum)},
                { Tuple.Create(DeclarationType.Enumeration, Accessibility.Private ), GetImageSource(resx.ObjectEnumPrivate)},
                { Tuple.Create(DeclarationType.EnumerationMember, Accessibility.Public), GetImageSource(resx.ObjectEnumItem)},
                { Tuple.Create(DeclarationType.Event, Accessibility.Public), GetImageSource(resx.ObjectEvent)},
                { Tuple.Create(DeclarationType.Event, Accessibility.Private ), GetImageSource(resx.ObjectEventPrivate)},
                { Tuple.Create(DeclarationType.Function, Accessibility.Public), GetImageSource(resx.ObjectMethod)},
                { Tuple.Create(DeclarationType.Function, Accessibility.Friend ), GetImageSource(resx.ObjectMethodFriend)},
                { Tuple.Create(DeclarationType.Function, Accessibility.Private ), GetImageSource(resx.ObjectMethodPrivate)},
                { Tuple.Create(DeclarationType.LibraryFunction, Accessibility.Public), GetImageSource(resx.ObjectMethodShortcut)},
                { Tuple.Create(DeclarationType.LibraryProcedure, Accessibility.Public), GetImageSource(resx.ObjectMethodShortcut)},
                { Tuple.Create(DeclarationType.LibraryFunction, Accessibility.Private), GetImageSource(resx.ObjectMethodShortcut)},
                { Tuple.Create(DeclarationType.LibraryProcedure, Accessibility.Private), GetImageSource(resx.ObjectMethodShortcut)},
                { Tuple.Create(DeclarationType.LibraryFunction, Accessibility.Friend), GetImageSource(resx.ObjectMethodShortcut)},
                { Tuple.Create(DeclarationType.LibraryProcedure, Accessibility.Friend), GetImageSource(resx.ObjectMethodShortcut)},
                { Tuple.Create(DeclarationType.Procedure, Accessibility.Public), GetImageSource(resx.ObjectMethod)},
                { Tuple.Create(DeclarationType.Procedure, Accessibility.Friend ), GetImageSource(resx.ObjectMethodFriend)},
                { Tuple.Create(DeclarationType.Procedure, Accessibility.Private ), GetImageSource(resx.ObjectMethodPrivate)},
                { Tuple.Create(DeclarationType.PropertyGet, Accessibility.Public), GetImageSource(resx.ObjectProperties)},
                { Tuple.Create(DeclarationType.PropertyGet, Accessibility.Friend ), GetImageSource(resx.ObjectPropertiesFriend)},
                { Tuple.Create(DeclarationType.PropertyGet, Accessibility.Private ), GetImageSource(resx.ObjectPropertiesPrivate)},
                { Tuple.Create(DeclarationType.PropertyLet, Accessibility.Public), GetImageSource(resx.ObjectProperties)},
                { Tuple.Create(DeclarationType.PropertyLet, Accessibility.Friend ), GetImageSource(resx.ObjectPropertiesFriend)},
                { Tuple.Create(DeclarationType.PropertyLet, Accessibility.Private ), GetImageSource(resx.ObjectPropertiesPrivate)},
                { Tuple.Create(DeclarationType.PropertySet, Accessibility.Public), GetImageSource(resx.ObjectProperties)},
                { Tuple.Create(DeclarationType.PropertySet, Accessibility.Friend ), GetImageSource(resx.ObjectPropertiesFriend)},
                { Tuple.Create(DeclarationType.PropertySet, Accessibility.Private ), GetImageSource(resx.ObjectPropertiesPrivate)},
                { Tuple.Create(DeclarationType.UserDefinedType, Accessibility.Public), GetImageSource(resx.ObjectValueType)},
                { Tuple.Create(DeclarationType.UserDefinedType, Accessibility.Private ), GetImageSource(resx.ObjectValueTypePrivate)},
                { Tuple.Create(DeclarationType.UserDefinedTypeMember, Accessibility.Public), GetImageSource(resx.ObjectField)},
                { Tuple.Create(DeclarationType.Variable, Accessibility.Private), GetImageSource(resx.ObjectFieldPrivate)},
                { Tuple.Create(DeclarationType.Variable, Accessibility.Public ), GetImageSource(resx.ObjectField)},
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

        private string RemoveExtraWhiteSpace(string value)
        {
            var newStr = new StringBuilder();
            var trimmedJoinedString = value.Replace(" _\r\n", " ").Trim();

            for (var i = 0; i < trimmedJoinedString.Length; i++)
            {
                // this will not throw because `Trim` ensures the first character is not whitespace
                if (char.IsWhiteSpace(trimmedJoinedString[i]) && char.IsWhiteSpace(trimmedJoinedString[i - 1]))
                {
                    continue;
                }

                newStr.Append(trimmedJoinedString[i]);
            }

            return newStr.ToString();
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
                    _signature = Name.Insert(Name.Length - 6, RemoveExtraWhiteSpace(context.GetText())); 
                }
                else
                {
                    _signature = Name + RemoveExtraWhiteSpace(context.GetText());
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
