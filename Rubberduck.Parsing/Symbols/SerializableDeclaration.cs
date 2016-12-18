using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    [DataContract]
    public class SerializableDeclarationTree
    {
        [DataMember(IsRequired = true)]
        public readonly SerializableDeclaration Node;

        [DataMember(IsRequired = true)]
        public IEnumerable<SerializableDeclarationTree> Children;

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

    [DataContract]
    public class SerializableMemberAttribute
    {
        public SerializableMemberAttribute(string name, IEnumerable<string> values)
        {
            Name = name;
            Values = values;
        }

        [DataMember(IsRequired = true)]
        public readonly string Name;

        [DataMember(IsRequired = true)]
        public readonly IEnumerable<string> Values;
    }

    public class SerializableDeclaration
    {
        public SerializableDeclaration()
        { }

        public SerializableDeclaration(Declaration declaration)
        {
            IdentifierName = declaration.IdentifierName;

            Attributes = declaration.Attributes
                .Select(a => new SerializableMemberAttribute(a.Key, a.Value))
                .ToList();

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

            var param = declaration as ParameterDeclaration;
            if (param != null)
            {
                IsOptionalParam = param.IsOptional;
                IsByRefParam = param.IsByRef;
                IsParamArray = param.IsParamArray;
            }
        }

        public List<SerializableMemberAttribute> Attributes { get; set; }

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

        public bool IsOptionalParam { get; set; }
        public bool IsByRefParam { get; set; }
        public bool IsParamArray { get; set; }

        public Declaration Unwrap(Declaration parent)
        {
            var annotations = Enumerable.Empty<IAnnotation>();
            var attributes = new Attributes();
            foreach (var attribute in Attributes)
            {
                attributes.Add(attribute.Name, attribute.Values);
            }

            switch (DeclarationType)
            {
                case DeclarationType.Project:
                    return new ProjectDeclaration(QualifiedMemberName, IdentifierName, true);
                case DeclarationType.ClassModule:
                    return new ClassModuleDeclaration(QualifiedMemberName, parent, IdentifierName, true, annotations, attributes);
                case DeclarationType.ProceduralModule:
                    return new ProceduralModuleDeclaration(QualifiedMemberName, parent, IdentifierName, true, annotations, attributes);
                case DeclarationType.Procedure:
                    return new SubroutineDeclaration(QualifiedMemberName, parent, parent, AsTypeName, Accessibility, null, Selection.Empty, true, annotations, attributes);
                case DeclarationType.Function:
                    return new FunctionDeclaration(QualifiedMemberName, parent, parent, AsTypeName, null, TypeHint, Accessibility, null, Selection.Empty, IsArray, true, annotations, attributes);
                case DeclarationType.PropertyGet:
                    return new PropertyGetDeclaration(QualifiedMemberName, parent, parent, AsTypeName, null, TypeHint, Accessibility, null, Selection.Empty, IsArray, true, annotations, attributes);
                case DeclarationType.PropertyLet:
                    return new PropertyLetDeclaration(QualifiedMemberName, parent, parent, AsTypeName, Accessibility, null, Selection.Empty, true, annotations, attributes);
                case DeclarationType.PropertySet:
                    return new PropertySetDeclaration(QualifiedMemberName, parent, parent, AsTypeName, Accessibility, null, Selection.Empty, true, annotations, attributes);
                case DeclarationType.Parameter:
                    return new ParameterDeclaration(QualifiedMemberName, parent, AsTypeName, null, TypeHint, IsOptionalParam, IsByRefParam, IsArray, IsParamArray);

                default:
                    return new Declaration(QualifiedMemberName, parent, ParentScope, AsTypeName, TypeHint, IsSelfAssigned, IsWithEvents, Accessibility, DeclarationType, null, Selection.Empty, IsArray, null, IsBuiltIn, null, attributes);
            }
        }
    }
}