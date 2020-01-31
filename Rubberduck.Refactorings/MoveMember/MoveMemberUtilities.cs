using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using System;
using System.Collections.Generic;
using System.Linq;


namespace Rubberduck.Refactorings.MoveMember
{
    public struct MoveElementGroups
    {
        private IEnumerable<Declaration> _allDeclarations;
        public MoveElementGroups(IEnumerable<Declaration> allDeclarations)
        {
            _allDeclarations = allDeclarations.Where(ge => !ge.DeclarationType.HasFlag(DeclarationType.Parameter));
        }

        public IEnumerable<Declaration> AllDeclarations => _allDeclarations;
        public IEnumerable<Declaration> Private => _allDeclarations.Where(p => p.HasPrivateAccessibility());
        public IEnumerable<Declaration> Public => _allDeclarations.Except(Private);
        public IEnumerable<Declaration> Members => _allDeclarations.Where(p => p.IsMember());
        public IEnumerable<Declaration> Types => _allDeclarations.Except(Members).Where(m => IsTypeDeclaration(m));
        public IEnumerable<Declaration> NonMembers => _allDeclarations.Except(Members).Except(Types);
        public IEnumerable<Declaration> PrivateMembers => Private.Where(p => p.IsMember());
        public IEnumerable<Declaration> PublicMembers => Public.Where(p => p.IsMember());
        public IEnumerable<Declaration> PrivateNonMembers => Private.Where(p => !p.IsMember() && !IsTypeDeclaration(p));
        public IEnumerable<Declaration> PublicNonMembers => Public.Where(p => !p.IsMember() && !IsTypeDeclaration(p));
        public IEnumerable<Declaration> PrivateTypeDefinitions => Private.Where(p => IsTypeDeclaration(p));
        public IEnumerable<Declaration> PublicTypeDefinitions => Public.Where(p => IsTypeDeclaration(p));

        public bool Contains(Declaration declaration) => AllDeclarations.Contains(declaration);
        public IEnumerable<Declaration> Concat(MoveElementGroups group) => AllDeclarations.Concat(group.AllDeclarations);
        public IEnumerable<Declaration> Concat(IEnumerable<Declaration> declarations) => AllDeclarations.Concat(declarations);
        public Declaration FirstOrDefault() => AllDeclarations.FirstOrDefault();
        public Declaration First() => AllDeclarations.First();
        public Declaration SingleOrDefault() => AllDeclarations.SingleOrDefault();
        public Declaration Single() => AllDeclarations.Single();
        public IEnumerable<IdentifierReference> AllReferences() => AllDeclarations.AllReferences();
        public IEnumerable<Declaration> Except(IEnumerable<Declaration> exceptions) => AllDeclarations.Except(exceptions);
        public IEnumerable<Declaration> Except(MoveElementGroups exceptionGroups) => AllDeclarations.Except(exceptionGroups.AllDeclarations);

        private static bool IsTypeDeclaration(Declaration declaration)
        {
            return declaration.DeclarationType.HasFlag(DeclarationType.Enumeration) 
                || declaration.DeclarationType.HasFlag(DeclarationType.UserDefinedType);
        }
    }

    public struct PropertyBlockProvider
    {
        public PropertyBlockProvider(Declaration prototype, string backingVariableIdentifier, string propertyName = null, string letSetArgIdentifier = "value")
        {
            Identifier = propertyName ?? prototype.IdentifierName;
            AsTypeName = prototype.AsTypeName;
            BackingVariableIdentifier = backingVariableIdentifier;
            LetSetArgIdentifier = letSetArgIdentifier;
        }

        public string Identifier;
        public string AsTypeName;
        public string BackingVariableIdentifier;
        public string LetSetArgIdentifier;

        public string PropertyLet =>
$@"{Tokens.Public} {Tokens.Property} {Tokens.Let} {Identifier}({Tokens.ByVal} {LetSetArgIdentifier} {Tokens.As} {AsTypeName})
    {BackingVariableIdentifier} = {LetSetArgIdentifier}
{Tokens.End} {Tokens.Property}
";

        public string PropertySet =>
$@"{Tokens.Public} {Tokens.Property} {Tokens.Set} {Identifier}({Tokens.ByRef} {LetSetArgIdentifier} {Tokens.As} {AsTypeName})
    {Tokens.Set} {BackingVariableIdentifier} = {LetSetArgIdentifier}
{Tokens.End} {Tokens.Property}
";

        public string PropertyGet =>
$@"{Tokens.Public} {Tokens.Property} {Tokens.Get} {Identifier}() {Tokens.As} {AsTypeName}
    {Identifier} = {BackingVariableIdentifier}
{Tokens.End} {Tokens.Property}
";

        public string BackingVariableDeclaration =>
           $"{Tokens.Private} {BackingVariableIdentifier} {Tokens.As} {AsTypeName}";
    }

    public struct ClassVariableContentProvider
    {
        private readonly string _classModuleName;
        private readonly string _classVariableName;

        public ClassVariableContentProvider(string classModuleName, string classVariableName = null)
        {
            _classModuleName = classModuleName;
            _classVariableName = classVariableName ?? $"{MoveMemberResources.Prefix_Variable}{_classModuleName}";
        }

        public string ClassVariableIdentifier => _classVariableName;

        public string ClassVariableDeclaration =>
              $"{Tokens.Private} {ClassVariableIdentifier} {Tokens.As} {_classModuleName}";

        public string ClassInstantiationFragment =>
            $"{Tokens.Set} {ClassVariableIdentifier} = {Tokens.New} {_classModuleName}";

        public string ClassModuleClassInitializeProcedure =>
            $@"
{PRIVATE_SUB} {MoveMemberResources.Class_Initialize}()
    {ClassInstantiationFragment}
{END_SUB}
";
        public string StdModuleClassVariableInstantiationSubName =>
            $"{MoveMemberResources.Prefix_ClassInstantiationProcedure}{ClassVariableIdentifier}";

        public string StdModuleClassVariableInstantiationProcedure =>
$@"
{PRIVATE_SUB} {$"{StdModuleClassVariableInstantiationSubName}"}()
    {Tokens.If} {ClassVariableIdentifier} {Tokens.Is} {Tokens.Nothing} {Tokens.Then}
        {ClassInstantiationFragment}
    {END_IF}
{END_SUB}";

        private string END_IF => $"{Tokens.End} {Tokens.If}";
        private string PRIVATE_SUB => $"{Tokens.Private} {Tokens.Sub}";
        private string END_SUB => $"{Tokens.End} {Tokens.Sub}";
    }
}
