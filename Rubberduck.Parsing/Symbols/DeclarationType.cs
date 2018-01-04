using System;
using System.Collections.Generic;

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
            Name = name;
        }

        public string Name { get; }
    }

    public class ProjectNode : IIdentifier
    {
        public ProjectNode(IdentifierNode identifier, IReadOnlyList<ModuleNode> modules)
        {
            Identifier = identifier;
            Modules = modules;
        }

        public IdentifierNode Identifier { get; }

        public IReadOnlyList<ModuleNode> Modules { get; }
    }

    public class ModuleNode : IIdentifier
    {
        public ModuleNode(IdentifierNode identifier, IReadOnlyList<ModuleOptionNode> options, IReadOnlyList<MemberNode> members)
        {
            Identifier = identifier;
            Options = options;
            Members = members;
        }

        public IdentifierNode Identifier { get; }

        public IReadOnlyList<ModuleOptionNode> Options { get; }

        public IReadOnlyList<MemberNode> Members { get; }
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
            Value = value;
        }

        public int Value { get; }
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
            Value = value;
        }

        public OptionCompareEnum Value { get; }
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
            Identifier = identifier;
        }

        public IdentifierNode Identifier { get; }
    }

    public class EnumerationNode : NamedMemberNode
    {
        public EnumerationNode(IdentifierNode identifier, IReadOnlyList<EnumerationMemberNode> members) 
            : base(identifier)
        {
            Members = members;
        }

        public IReadOnlyList<EnumerationMemberNode> Members { get; }
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
            Members = members;
        }

        public IReadOnlyList<UserDefinedTypeMemberNode> Members { get; }
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
            Members = members;
        }

        public IReadOnlyList<NamedMemberNode> Members { get; }
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
            Identifier = identifier;
            Ordinal = ordinal;
        }

        public IdentifierNode Identifier { get; }

        public int Ordinal { get; }
    }

    public class ProcedureMemberNode : NamedMemberNode
    {
        public ProcedureMemberNode(IdentifierNode identifier, IReadOnlyList<ParameterDeclarationNode> parameters) 
            : base(identifier)
        {
            Parameters = parameters;
        }

        public IReadOnlyList<ParameterDeclarationNode> Parameters { get; }
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
