using System;
using System.Collections.Generic;

namespace Rubberduck.Parsing.Symbols
{
    [Flags]
    public enum DeclarationType : long
    {
        Project = 1L << 0,
        Module = 1L << 1,
        ProceduralModule = 1L << 2 | Module,
        ClassModule = 1L << 3 | Module,
        UserForm = 1L << 4 | ClassModule,
        Document = 1L << 5 | ClassModule,
        VbForm = 1L << 6 | ClassModule,
        Member = 1L << 7,
        Procedure = 1L << 8 | Member,
        Function = 1L << 9 | Member,
        Property = 1L << 10 | Member,
        PropertyGet = 1L << 11 | Property | Function,
        PropertyLet = 1L << 12 | Property | Procedure,
        PropertySet = 1L << 13 | Property | Procedure,
        Parameter = 1L << 14,
        Variable = 1L << 15,
        Control = 1L << 16 | Variable,
        Constant = 1L << 17,
        Enumeration = 1L << 18,
        EnumerationMember = 1L << 19,
        Event = 1L << 20,
        UserDefinedType = 1L << 21,
        UserDefinedTypeMember = 1L << 22,
        LibraryFunction = 1L << 23 | Function,
        LibraryProcedure = 1L << 24 | Procedure,
        LineLabel = 1L << 25,
        UnresolvedMember = 1L << 26,
        BracketedExpression = 1L << 27,
        ComAlias = 1L << 28,
        MdiForm = 1L << 29 | VbForm,
        ResFile = 1L << 30,
        PropPage = 1L << 31 | ClassModule,
        UserControl = 1L << 32 | ClassModule,
        DocObject = 1L << 33 | ClassModule,
        RelatedDocument = 1L << 34,
        ActiveXDesigner = 1L << 35 | ClassModule
    }

    public interface IIdentifier { IdentifierNode Identifier { get; } }

    public class IdentifierNode
    {
        public IdentifierNode(string name)
        {
            _name = name;
        }

        private readonly string _name;
        public string Name { get { return _name; } }
    }

    public class ProjectNode : IIdentifier
    {
        public ProjectNode(IdentifierNode identifier, IReadOnlyList<ModuleNode> modules)
        {
            _identifier = identifier;
            _modules = modules;
        }

        private readonly IdentifierNode _identifier;
        private readonly IReadOnlyList<ModuleNode> _modules;

        public IdentifierNode Identifier { get { return _identifier; } }
        public IReadOnlyList<ModuleNode> Modules { get { return _modules; } } 
    }

    public class ModuleNode : IIdentifier
    {
        public ModuleNode(IdentifierNode identifier, IReadOnlyList<ModuleOptionNode> options, IReadOnlyList<MemberNode> members)
        {
            _identifier = identifier;
            _options = options;
            _members = members;
        }

        private readonly IdentifierNode _identifier;
        public IdentifierNode Identifier { get { return _identifier; } }

        private readonly IReadOnlyList<ModuleOptionNode> _options;
        public IReadOnlyList<ModuleOptionNode> Options { get { return _options; } } 

        private readonly IReadOnlyList<MemberNode> _members;
        public IReadOnlyList<MemberNode> Members { get { return _members; } }
    }

    public class ClassModuleNode : ModuleNode
    {
        public enum ClassFlags
        {
            GlobalNamespace,
            Creatable,
            PredeclaredId,
            Exposed
        }

        public ClassModuleNode(IdentifierNode identifier, IReadOnlyList<ModuleOptionNode> options, IReadOnlyList<MemberNode> members) 
            : base(identifier, options, members)
        {
        }

        // expose attributes?
    }

    public abstract class ModuleOptionNode
    {
    }

    public class OptionExplicitNode : ModuleOptionNode
    {
    }

    public class OptionPrivateModuleNode : ModuleOptionNode
    {
    }

    public class OptionBaseNode : ModuleOptionNode
    {
        public OptionBaseNode(int value)
        {
            _value = value;
        }

        private readonly int _value;
        public int Value { get { return _value; } }
    }

    public class OptionCompareNode : ModuleOptionNode
    {
        public enum OptionCompareEnum
        {
            Binary,
            Text,
            Database
        }

        public OptionCompareNode(OptionCompareEnum value)
        {
            _value = value;
        }

        private readonly OptionCompareEnum _value;
        public OptionCompareEnum Value { get { return _value; } }
    }

    /// <summary>
    /// Base class for a node whose immediate parent is a module.
    /// </summary>
    public abstract class MemberNode
    {
    }

    public abstract class NamedMemberNode : MemberNode, IIdentifier
    {
        protected NamedMemberNode(IdentifierNode identifier)
        {
            _identifier = identifier;
        }

        private readonly IdentifierNode _identifier;
        public IdentifierNode Identifier { get { return _identifier; } }
    }

    public class EnumerationNode : NamedMemberNode
    {
        public EnumerationNode(IdentifierNode identifier, IReadOnlyList<EnumerationMemberNode> members) 
            : base(identifier)
        {
            _members = members;
        }

        private readonly IReadOnlyList<EnumerationMemberNode> _members;
        public IReadOnlyList<EnumerationMemberNode> Members { get { return _members; } } 
    }

    public class EnumerationMemberNode : NamedMemberNode
    {
        public EnumerationMemberNode(IdentifierNode identifier)
            : base(identifier)
        {
        }
    }

    public class UserDefinedTypeNode : NamedMemberNode
    {
        public UserDefinedTypeNode(IdentifierNode identifier, IReadOnlyList<UserDefinedTypeMemberNode> members) 
            : base(identifier)
        {
            _members = members;
        }

        private readonly IReadOnlyList<UserDefinedTypeMemberNode> _members;
        public IReadOnlyList<UserDefinedTypeMemberNode> Members { get { return _members; } }
    }

    public class UserDefinedTypeMemberNode : NamedMemberNode
    {
        public UserDefinedTypeMemberNode(IdentifierNode identifier) 
            : base(identifier)
        {
        }
    }

    public class DeclarationStatementNode : MemberNode
    {
        public DeclarationStatementNode(IReadOnlyList<NamedMemberNode> members)
        {
            _members = members;
        }

        private readonly IReadOnlyList<NamedMemberNode> _members;
        public IReadOnlyList<NamedMemberNode> Members { get { return _members; } }
    }

    public class VariableDeclarationNode : NamedMemberNode
    {
        public VariableDeclarationNode(IdentifierNode identifier)
            : base(identifier)
        {
        }
    }

    public class ConstantDeclarationNode : NamedMemberNode
    {
        public ConstantDeclarationNode(IdentifierNode identifier)
            : base(identifier)
        {
        }
    }

    public class ArrayDeclarationNode : VariableDeclarationNode
    {
        public ArrayDeclarationNode(IdentifierNode identifier)
            : base(identifier)
        {
        }
    }

    public class ParameterDeclarationNode : IIdentifier
    {
        public ParameterDeclarationNode(IdentifierNode identifier, int ordinal)
        {
            _identifier = identifier;
            _ordinal = ordinal;
        }

        private readonly IdentifierNode _identifier;
        public IdentifierNode Identifier { get { return _identifier; } }

        private readonly int _ordinal;
        public int Ordinal { get { return _ordinal; } }
    }

    public class ProcedureMemberNode : NamedMemberNode
    {
        public ProcedureMemberNode(IdentifierNode identifier, IReadOnlyList<ParameterDeclarationNode> parameters) 
            : base(identifier)
        {
            _parameters = parameters;
        }

        private readonly IReadOnlyList<ParameterDeclarationNode> _parameters;
        public IReadOnlyList<ParameterDeclarationNode> Parameters { get { return _parameters; } }
    }

    public class FunctionMemberNode : ProcedureMemberNode
    {
        public FunctionMemberNode(IdentifierNode identifier, IReadOnlyList<ParameterDeclarationNode> parameters) 
            : base(identifier, parameters)
        {
        }
    }

    public class PropertyMemberNode : ProcedureMemberNode
    {
        public PropertyMemberNode(IdentifierNode identifier, IReadOnlyList<ParameterDeclarationNode> parameters) 
            : base(identifier, parameters)
        {
        }
    }

    public class PropertyGetMemberNode : PropertyMemberNode
    {
        public PropertyGetMemberNode(IdentifierNode identifier, IReadOnlyList<ParameterDeclarationNode> parameters) 
            : base(identifier, parameters)
        {
        }
    }

    public class PropertyLetMemberNode : PropertyMemberNode
    {
        public PropertyLetMemberNode(IdentifierNode identifier, IReadOnlyList<ParameterDeclarationNode> parameters) 
            : base(identifier, parameters)
        {
        }
    }

    public class PropertySetMemberNode : PropertyMemberNode
    {
        public PropertySetMemberNode(IdentifierNode identifier, IReadOnlyList<ParameterDeclarationNode> parameters) 
            : base(identifier, parameters)
        {
        }
    }

    public class DeclareStatementMemberNode : NamedMemberNode
    {
        public DeclareStatementMemberNode(IdentifierNode identifier) 
            : base(identifier)
        {
        }
    }
}
