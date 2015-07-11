using Antlr4.Runtime;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Parsing.Symbols
{
    public class ValuedDeclaration : Declaration
    {
        /// <summary>
        /// Creates a new valued built-in declaration.
        /// </summary>
        public ValuedDeclaration(QualifiedMemberName qualifiedName, string parentScope,
            string asTypeName, Accessibility accessibility, DeclarationType declarationType, string value, IRubberduckFactory<IRubberduckCodePane> factory)
            : this(qualifiedName, parentScope, asTypeName, accessibility, declarationType, value, null, Selection.Home, factory, true)
        {
        }

        public ValuedDeclaration(QualifiedMemberName qualifiedName, string parentScope,
            string asTypeName, Accessibility accessibility, DeclarationType declarationType, string value,
            ParserRuleContext context, Selection selection, IRubberduckFactory<IRubberduckCodePane> factory, bool isBuiltIn = false)
            :base(qualifiedName, parentScope, asTypeName, true, false, accessibility, declarationType, context, selection, factory, isBuiltIn)
        {
            _value = value;
        }

        private readonly string _value;
        /// <summary>
        /// Gets a string representing the declared value.
        /// </summary>
        public string Value { get { return _value; } }
    }
}