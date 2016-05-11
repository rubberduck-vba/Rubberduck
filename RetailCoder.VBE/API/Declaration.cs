using System.Collections.Generic;
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
        public DeclarationType DeclarationType {get { return TypeMappings[_declaration.DeclarationType]; }}
        public string TypeName { get { return _declaration.AsTypeName; } }
        public bool IsArray { get { return _declaration.IsArray(); } }

        private static readonly IDictionary<Parsing.Symbols.DeclarationType,DeclarationType> TypeMappings =
            new Dictionary<Parsing.Symbols.DeclarationType, DeclarationType>
            {
                { Parsing.Symbols.DeclarationType.Project, DeclarationType.Project },
                { Parsing.Symbols.DeclarationType.ProceduralModule, DeclarationType.StandardModule },
                { Parsing.Symbols.DeclarationType.ClassModule, DeclarationType.ClassModule },
                { Parsing.Symbols.DeclarationType.Control, DeclarationType.Control },
                { Parsing.Symbols.DeclarationType.UserForm, DeclarationType.UserForm },
                { Parsing.Symbols.DeclarationType.Document, DeclarationType.Document },
                { Parsing.Symbols.DeclarationType.ModuleOption, DeclarationType.ModuleOption },
                { Parsing.Symbols.DeclarationType.Procedure, DeclarationType.Procedure },
                { Parsing.Symbols.DeclarationType.Function, DeclarationType.Function },
                { Parsing.Symbols.DeclarationType.PropertyGet, DeclarationType.PropertyGet },
                { Parsing.Symbols.DeclarationType.PropertyLet, DeclarationType.PropertyLet },
                { Parsing.Symbols.DeclarationType.PropertySet, DeclarationType.PropertySet },
                { Parsing.Symbols.DeclarationType.Parameter, DeclarationType.Parameter },
                { Parsing.Symbols.DeclarationType.Variable, DeclarationType.Variable },
                { Parsing.Symbols.DeclarationType.Constant, DeclarationType.Constant },
                { Parsing.Symbols.DeclarationType.Enumeration, DeclarationType.Enumeration },
                { Parsing.Symbols.DeclarationType.EnumerationMember, DeclarationType.EnumerationMember },
                { Parsing.Symbols.DeclarationType.Event, DeclarationType.Event },
                { Parsing.Symbols.DeclarationType.UserDefinedType, DeclarationType.UserDefinedType },
                { Parsing.Symbols.DeclarationType.UserDefinedTypeMember, DeclarationType.UserDefinedTypeMember },
                { Parsing.Symbols.DeclarationType.LibraryFunction, DeclarationType.LibraryFunction },
                { Parsing.Symbols.DeclarationType.LibraryProcedure, DeclarationType.LibraryProcedure },
                { Parsing.Symbols.DeclarationType.LineLabel, DeclarationType.LineLabel },
            };

        private Declaration _parentDeclaration;
        public Declaration ParentDeclaration
        {
            get
            {
                return _parentDeclaration ?? (_parentDeclaration = new Declaration(Instance.ParentDeclaration));
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
