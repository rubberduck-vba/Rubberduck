using System.Collections.Generic;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    public class SerializableDeclaration
    {
        public SerializableDeclaration()
        { }

        public SerializableDeclaration(Declaration declaration)
        {
            ParentDeclaration = declaration.ParentDeclaration == null ? null : new SerializableDeclaration(declaration.ParentDeclaration);
            ParentScope = declaration.ParentScopeDeclaration == null ? null : new SerializableDeclaration(declaration.ParentScopeDeclaration);
            AsTypeDeclaration = declaration.AsTypeDeclaration == null ? null : new SerializableDeclaration(declaration.AsTypeDeclaration);
            QualifiedMemberName = declaration.QualifiedName;
            Annotations = declaration.Annotations;
            TypeHint = declaration.TypeHint;
            AsTypeName = declaration.AsTypeName;
            IsArray = declaration.IsArray;
            IsBuiltIn = declaration.IsBuiltIn;
            IsSelfAssigned = declaration.IsSelfAssigned;
            IsWithEvents = declaration.IsWithEvents;
            Accessibility = declaration.Accessibility;
            DeclarationType = declaration.DeclarationType;
        }

        public QualifiedMemberName QualifiedMemberName { get; set; }
        public IEnumerable<IAnnotation> Annotations { get; set; }
        public Attributes Attributes { get; set; }

        public SerializableDeclaration ParentDeclaration { get; set; }
        public SerializableDeclaration ParentScope { get; set; }
        public SerializableDeclaration AsTypeDeclaration { get; set; }
        public string AsTypeName { get; set; }
        public string TypeHint { get; set; }
        public bool IsArray { get; set; }
        public bool IsBuiltIn { get; set; }
        public bool IsSelfAssigned { get; set; }
        public bool IsWithEvents { get; set; }
        public Accessibility Accessibility { get; set; }
        public DeclarationType DeclarationType { get; set; }

        public Declaration Unwrap()
        {
            return new Declaration(QualifiedMemberName, ParentDeclaration.Unwrap(), ParentScope.Unwrap(), AsTypeName, TypeHint, IsSelfAssigned, IsWithEvents, Accessibility, DeclarationType, null, Selection.Empty, IsArray, null, IsBuiltIn, Annotations, Attributes);
        }
    }
}