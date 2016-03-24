using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using RubberduckDeclaration = Rubberduck.Parsing.Symbols.Declaration;

namespace Rubberduck.API
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDeclaration
    {
        string Name { get; }
        Accessibility Accessibility { get; }
        DeclarationType DeclarationType { get; }
        string TypeName { get; }
        bool IsArray { get; }

        Declaration ParentDeclaration { get; }

        IdentifierReference[] References { get; }
    }

    [ComVisible(true)]
    [Guid(ClassId)]
    [ProgId(ProgId)]
    [ComDefaultInterface(typeof(IDeclaration))]
    [EditorBrowsable(EditorBrowsableState.Always)]
    public class Declaration : IDeclaration
    {
        private const string ClassId = "67940D0B-081A-45BE-B0B9-CAEAFE034BC0";
        private const string ProgId = "Rubberduck.Declaration";

        private readonly RubberduckDeclaration _declaration;

        internal Declaration(RubberduckDeclaration declaration)
        {
            _declaration = declaration;
        }

        protected RubberduckDeclaration Instance { get { return _declaration; } }

        public string Name { get { return _declaration.IdentifierName; } }
        public Accessibility Accessibility { get { return (Accessibility)_declaration.Accessibility; } }
        public DeclarationType DeclarationType {get { return (DeclarationType)_declaration.DeclarationType; }}
        public string TypeName { get { return _declaration.AsTypeName; } }
        public bool IsArray { get { return _declaration.IsArray(); } }

        private Declaration _parentDeclaration;
        public Declaration ParentDeclaration
        {
            get
            {
                return _parentDeclaration ?? (_parentDeclaration = new Declaration(Instance));
            }
        }

        private IdentifierReference[] _references;
        public IdentifierReference[] References
        {
            get
            {
                return _references ?? (_references = _declaration.References.Select(item => new IdentifierReference(item)).ToArray());
            }
        }
    }
}
