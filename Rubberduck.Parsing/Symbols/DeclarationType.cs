using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace Rubberduck.Parsing.Symbols
{
    [Flags]
    public enum DeclarationType
    {
        Project = 1 << 0,
        Module = 1 << 1,
        ProceduralModule = 1 << 2 | Module,
        ClassModule = 1 << 3 | Module,
        UserForm = 1 << 4 | ClassModule,
        Document = 1 << 5 | ClassModule,
        ModuleOption = 1 << 6,
        Member = 1 << 7,
        Procedure = 1 << 8 | Member,
        Function = 1 << 9 | Member,
        Property = 1 << 10 | Member,
        PropertyGet = 1 << 11 | Property | Function,
        PropertyLet = 1 << 12 | Property | Procedure,
        PropertySet = 1 << 13 | Property | Procedure,
        Parameter = 1 << 14,
        Variable = 1 << 15,
        Control = 1 << 16 | Variable,
        Constant = 1 << 17,
        Enumeration = 1 << 18,
        EnumerationMember = 1 << 19,
        Event = 1 << 20,
        UserDefinedType = 1 << 21,
        UserDefinedTypeMember = 1 << 22,
        LibraryFunction = 1 << 23 | Function,
        LibraryProcedure = 1 << 24 | Procedure,
        LineLabel = 1 << 25,
        UnresolvedMember = 1 << 26,
        BracketedExpression = 1 << 27,
        ComAlias = 1 << 28
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
