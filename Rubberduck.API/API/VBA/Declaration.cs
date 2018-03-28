﻿using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using RubberduckDeclaration = Rubberduck.Parsing.Symbols.Declaration;

namespace Rubberduck.API.VBA
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDeclaration
    {
        [ComVisible(true)]
        string Name { get; }
        [ComVisible(true)]
        Accessibility Accessibility { get; }
        [ComVisible(true)]
        DeclarationType DeclarationType { get; }
        string TypeName { get; }
        [ComVisible(true)]
        bool IsArray { get; }
        [ComVisible(true)]
        Declaration ParentDeclaration { get; }
        [ComVisible(true)]
        IdentifierReference[] References { get; }
    }

    [ComVisible(true)]
    [Guid(RubberduckGuid.DeclarationClassGuid)]
    [ProgId(RubberduckProgId.DeclarationProgId)]
    [ComDefaultInterface(typeof(IDeclaration))]
    [EditorBrowsable(EditorBrowsableState.Always)]
    public class Declaration : IDeclaration
    {
        internal Declaration(RubberduckDeclaration declaration)
        {
            Instance = declaration;
        }

        protected RubberduckDeclaration Instance { get; }

        public bool IsArray => Instance.IsArray;
        public string Name => Instance.IdentifierName;
        public Accessibility Accessibility => (Accessibility)Instance.Accessibility;
        public DeclarationType DeclarationType => TypeMappings[Instance.DeclarationType];
        public string TypeName => Instance.AsTypeName;

        private static readonly IDictionary<Parsing.Symbols.DeclarationType,DeclarationType> TypeMappings =
            new Dictionary<Parsing.Symbols.DeclarationType, DeclarationType>
            {
                { Parsing.Symbols.DeclarationType.Project, DeclarationType.Project },
                { Parsing.Symbols.DeclarationType.ProceduralModule, DeclarationType.StandardModule },
                { Parsing.Symbols.DeclarationType.ClassModule, DeclarationType.ClassModule },
                { Parsing.Symbols.DeclarationType.Control, DeclarationType.Control },
                { Parsing.Symbols.DeclarationType.UserForm, DeclarationType.UserForm },
                { Parsing.Symbols.DeclarationType.Document, DeclarationType.Document },
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
        public Declaration ParentDeclaration => _parentDeclaration ?? (_parentDeclaration = new Declaration(Instance.ParentDeclaration));

        private IdentifierReference[] _references;
        public IdentifierReference[] References
        {
            get
            {
                return _references ?? (_references = Instance.References.Select(item => new IdentifierReference(item)).ToArray());
            }
        }
    }
}
