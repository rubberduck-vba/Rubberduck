using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.ComReflection
{
    public class SerializableDeclarationTree
    {
        public SerializableDeclaration Node;

        private List<SerializableDeclarationTree> _children = new List<SerializableDeclarationTree>();

        public IEnumerable<SerializableDeclarationTree> Children
        {
            get => _children;
            set => _children = new List<SerializableDeclarationTree>(value);
        }

        public SerializableDeclarationTree() { }

        public SerializableDeclarationTree(Declaration declaration)
            : this(new SerializableDeclaration(declaration)) { }

        public SerializableDeclarationTree(Declaration declaration, IEnumerable<SerializableDeclarationTree> children)
            : this(new SerializableDeclaration(declaration), children) { }

        public SerializableDeclarationTree(SerializableDeclaration node)
            : this(node, Enumerable.Empty<SerializableDeclarationTree>()) { }

        public SerializableDeclarationTree(SerializableDeclaration node, IEnumerable<SerializableDeclarationTree> children)
        {
            Node = node;
            Children = children;
        }

        public void SortChildren()
        {
            _children = _children.OrderBy(child => child.Node.DeclarationType).ThenBy(child => child.Node.IdentifierName).ToList();
            foreach (var child in _children)
            {
                child.SortChildren();
            }
        }

        public void AddChildren(IEnumerable<Declaration> declarations)
        {
            foreach (var child in declarations)
            {
                _children.Add(new SerializableDeclarationTree(child));
            }
        }

        public void AddChildTree(SerializableDeclarationTree tree)
        {
            _children.Add(tree);
        }

        public void AddChildTrees(IEnumerable<SerializableDeclarationTree> trees)
        {
            _children.AddRange(trees);
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

    [DataContract]
    public class SerializableProject
    {
        public SerializableProject() { }

        public SerializableProject(ProjectDeclaration declaration)
        {
            Node = new SerializableDeclaration(declaration);
            var project = (ProjectDeclaration)declaration;
            MajorVersion = project.MajorVersion;
            MinorVersion = project.MinorVersion;
        }

        [DataMember(IsRequired = true)]
        public SerializableDeclaration Node { get; set; }
        [DataMember(IsRequired = true)]

        private List<SerializableDeclarationTree> _declarations = new List<SerializableDeclarationTree>();

        public IEnumerable<SerializableDeclarationTree> Declarations
        {
            get => _declarations;
            set => _declarations = new List<SerializableDeclarationTree>(value);
        }

        [DataMember(IsRequired = true)]
        public long MajorVersion { get; set; }
        [DataMember(IsRequired = true)]
        public long MinorVersion { get; set; }

        public void AddDeclaration(SerializableDeclarationTree tree)
        {
            _declarations.Add(tree);
        }

        public void SortDeclarations()
        {
            _declarations = _declarations.OrderBy(declarationTree => declarationTree.Node.DeclarationType).ThenBy(declarationTree => declarationTree.Node.IdentifierName).ToList();
            foreach (var declarationTree in _declarations)
            {
                declarationTree.SortChildren();
            }
        }

        public List<Declaration> Unwrap()
        {
            var project = (ProjectDeclaration)Node.Unwrap(null);
            project.MajorVersion = MajorVersion;
            project.MinorVersion = MinorVersion;
            var output = new List<Declaration> { project };
            foreach (var declaration in Declarations)
            {
                output.AddRange(UnwrapTree(declaration, project));
            }
            return output;
        }

        private IEnumerable<Declaration> UnwrapTree(SerializableDeclarationTree tree, Declaration parent = null)
        {
            var current = tree.Node.Unwrap(parent);

            if (parent is ClassModuleDeclaration classModule && current.Attributes.HasDefaultMemberAttribute())
            {
                classModule.DefaultMember = current;
            }

            yield return current;

            foreach (var serializableDeclarationTree in tree.Children)
            {
                var unwrapped = UnwrapTree(serializableDeclarationTree, current);
                foreach (var declaration in unwrapped)
                {
                    yield return declaration;
                }
            }
        }
    }

    public class SerializableDeclaration
    {
        public SerializableDeclaration()
        { }

        public SerializableDeclaration(Declaration declaration)
        {
            IdentifierName = declaration.IdentifierName;

            Attributes = declaration.Attributes
                .Select(a => new SerializableMemberAttribute(a.Name, a.Values))
                .ToList();

            ParentScope = declaration.ParentScope;
            TypeHint = declaration.TypeHint;
            AsTypeName = declaration.AsTypeName;
            IsArray = declaration.IsArray;
            IsUserDefined = declaration.IsUserDefined;
            IsSelfAssigned = declaration.IsSelfAssigned;
            IsWithEvents = declaration.IsWithEvents;
            Accessibility = declaration.Accessibility;
            DeclarationType = declaration.DeclarationType;

            MemberName = declaration.QualifiedName.MemberName;
            ProjectName = declaration.QualifiedName.QualifiedModuleName.ProjectName;
            ProjectPath = declaration.QualifiedName.QualifiedModuleName.ProjectPath;
            ComponentName = declaration.QualifiedName.QualifiedModuleName.ComponentName;

            switch (declaration)
            {
                case ParameterDeclaration param:
                    IsOptionalParam = param.IsOptional;
                    IsByRefParam = param.IsByRef;
                    IsParamArray = param.IsParamArray;
                    DefaultValue = param.DefaultValue;
                    break;
                case ValuedDeclaration constant:
                    Expression = constant.Expression;
                    break;
                case ClassModuleDeclaration coclass:
                    IsControl = coclass.IsControl;
                    break;
            }
        }

        public List<SerializableMemberAttribute> Attributes { get; set; }

        public string IdentifierName { get; set; }

        public string MemberName { get; set; }
        public string ProjectName { get; set; }
        public string ProjectPath { get; set; }
        public string ComponentName { get; set; }

        public QualifiedModuleName QualifiedModuleName => new QualifiedModuleName(ProjectName, ProjectPath, ComponentName);
        public QualifiedMemberName QualifiedMemberName => new QualifiedMemberName(QualifiedModuleName, MemberName);

        public string ParentScope { get; set; }
        public string AsTypeName { get; set; }
        public string TypeHint { get; set; }
        public bool IsArray { get; set; }
        public bool IsUserDefined { get; set; }
        public bool IsSelfAssigned { get; set; }
        public bool IsWithEvents { get; set; }
        public bool IsExtensible { get; set; }
        public bool IsControl { get; set; }
        public Accessibility Accessibility { get; set; }
        public DeclarationType DeclarationType { get; set; }

        public bool IsOptionalParam { get; set; }
        public bool IsByRefParam { get; set; }
        public bool IsParamArray { get; set; }
        public string DefaultValue { get; set; }

        public string Expression { get; set; }

        public Declaration Unwrap(Declaration parent)
        {
            var annotations = Enumerable.Empty<IAnnotation>();
            var attributes = new Attributes();
            foreach (var attribute in Attributes)
            {
                attributes.Add(new AttributeNode(attribute.Name, attribute.Values));               
            }

            switch (DeclarationType)
            {
                case DeclarationType.Project:
                    return new ProjectDeclaration(QualifiedMemberName, IdentifierName, false, null);
                case DeclarationType.ClassModule:
                    return new ClassModuleDeclaration(QualifiedMemberName, parent, IdentifierName, false, annotations, attributes, false, IsControl);
                case DeclarationType.ProceduralModule:
                    return new ProceduralModuleDeclaration(QualifiedMemberName, parent, IdentifierName, false, annotations, attributes);
                case DeclarationType.Procedure:
                    return new SubroutineDeclaration(QualifiedMemberName, parent, parent, AsTypeName, Accessibility, null, null, Selection.Empty, false, annotations, attributes);
                case DeclarationType.Function:
                    return new FunctionDeclaration(QualifiedMemberName, parent, parent, AsTypeName, null, TypeHint, Accessibility, null, null, Selection.Empty, IsArray, false, annotations, attributes);
                case DeclarationType.Event:
                    return new EventDeclaration(QualifiedMemberName, parent, parent, AsTypeName, null, TypeHint, Accessibility, null, Selection.Empty, IsArray, false, annotations, attributes);
                case DeclarationType.PropertyGet:
                    return new PropertyGetDeclaration(QualifiedMemberName, parent, parent, AsTypeName, null, TypeHint, Accessibility, null, null, Selection.Empty, IsArray, false, annotations, attributes);
                case DeclarationType.PropertyLet:
                    return new PropertyLetDeclaration(QualifiedMemberName, parent, parent, AsTypeName, Accessibility, null, null, Selection.Empty, false, annotations, attributes);
                case DeclarationType.PropertySet:
                    return new PropertySetDeclaration(QualifiedMemberName, parent, parent, AsTypeName, Accessibility, null, null, Selection.Empty, false, annotations, attributes);
                case DeclarationType.Parameter:
                    var output = new ParameterDeclaration(QualifiedMemberName, parent, AsTypeName, null, TypeHint, IsOptionalParam, IsByRefParam, IsArray, IsParamArray, DefaultValue);
                    if (parent is IParameterizedDeclaration hasParams)
                    {
                        hasParams.AddParameter(output);
                    }
                    return output;
                case DeclarationType.EnumerationMember:
                    return new ValuedDeclaration(QualifiedMemberName, parent, ParentScope, AsTypeName, null, TypeHint, annotations, Accessibility, DeclarationType.EnumerationMember, Expression, null, Selection.Home, false);
                case DeclarationType.Constant:
                    return new ValuedDeclaration(QualifiedMemberName, parent, ParentScope, AsTypeName, null, TypeHint, annotations, Accessibility, DeclarationType.Constant, Expression, null, Selection.Home, false);
                default:
                    return new Declaration(QualifiedMemberName, parent, ParentScope, AsTypeName, TypeHint, IsSelfAssigned, IsWithEvents, Accessibility, DeclarationType, null, null, Selection.Empty, IsArray, null, IsUserDefined, null, attributes);
            }
        }
    }
}