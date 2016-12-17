using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    [DataContract]
    public class SerializableDeclarationTree
    {
        [DataMember(IsRequired = true)]
        public readonly SerializableDeclaration Node;

        [DataMember(IsRequired = true)]
        public readonly IEnumerable<SerializableDeclarationTree> Children;

        public SerializableDeclarationTree(Declaration declaration)   
            : this(new SerializableDeclaration(declaration)) { }

        public SerializableDeclarationTree(SerializableDeclaration node)
            : this(node, Enumerable.Empty<SerializableDeclarationTree>()) { }

        public SerializableDeclarationTree(SerializableDeclaration node, IEnumerable<SerializableDeclarationTree> children)
        {
            Node = node;
            Children = children;
        }
    }

    public class SerializableDeclaration
    {
        public SerializableDeclaration()
        { }

        public SerializableDeclaration(Declaration declaration)
        {
            IdentifierName = declaration.IdentifierName;

            //todo: figure these out
            //Annotations = declaration.Annotations.Cast<AnnotationBase>().ToArray();
            //Attributes = declaration.Attributes.ToDictionary(kvp => kvp.Key, kvp => kvp.Value.ToArray());

            ParentScope = declaration.ParentScope;
            TypeHint = declaration.TypeHint;
            AsTypeName = declaration.AsTypeName;
            IsArray = declaration.IsArray;
            IsBuiltIn = declaration.IsBuiltIn;
            IsSelfAssigned = declaration.IsSelfAssigned;
            IsWithEvents = declaration.IsWithEvents;
            Accessibility = declaration.Accessibility;
            DeclarationType = declaration.DeclarationType;

            MemberName = declaration.QualifiedName.MemberName;
            ProjectName = declaration.QualifiedName.QualifiedModuleName.ProjectName;
            ProjectPath = declaration.QualifiedName.QualifiedModuleName.ProjectPath;
            ComponentName = declaration.QualifiedName.QualifiedModuleName.ComponentName;
        }

        public string IdentifierName { get; set; }

        public string MemberName { get; set; }
        public string ProjectName { get; set; }
        public string ProjectPath { get; set; }
        public string ComponentName { get; set; }

        public QualifiedModuleName QualifiedModuleName { get { return new QualifiedModuleName(ProjectName, ProjectPath, ComponentName); } }
        public QualifiedMemberName QualifiedMemberName { get { return new QualifiedMemberName(QualifiedModuleName, MemberName); } }

        public string ParentScope { get; set; }
        public string AsTypeName { get; set; }
        public string TypeHint { get; set; }
        public bool IsArray { get; set; }
        public bool IsBuiltIn { get; set; }
        public bool IsSelfAssigned { get; set; }
        public bool IsWithEvents { get; set; }
        public Accessibility Accessibility { get; set; }
        public DeclarationType DeclarationType { get; set; }

        public Declaration Unwrap(Declaration parent)
        {
            return new Declaration(QualifiedMemberName, parent, ParentScope, AsTypeName, TypeHint, IsSelfAssigned, IsWithEvents, Accessibility, DeclarationType, null, Selection.Empty, IsArray, null, IsBuiltIn);
        }
    }
}