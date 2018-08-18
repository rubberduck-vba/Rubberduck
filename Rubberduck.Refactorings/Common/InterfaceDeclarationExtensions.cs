using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

// ReSharper disable once CheckNamespace
namespace Rubberduck.Common
{
    public static class InterfaceDeclarationExtensions
    {
        /// <summary>
        /// Returns the interface for a QualifiedSelection contained by a statement similar to "Implements IClass1"
        /// </summary>
        /// <param name="declarations"></param>
        /// <param name="selection"></param>
        /// <returns></returns>
        [SuppressMessage("ReSharper", "LoopCanBeConvertedToQuery")]
        public static Declaration FindInterface(this IEnumerable<Declaration> declarations, QualifiedSelection selection)
        {
            foreach (var declaration in declarations.FindInterfaces())
            {
                foreach (var reference in declaration.References)
                {
                    var implementsStmt = reference.Context.GetAncestor<VBAParser.ImplementsStmtContext>();

                    if (implementsStmt == null) { continue; }

                    var completeSelection = new Selection(implementsStmt.GetSelection().StartLine,
                        implementsStmt.GetSelection().StartColumn, reference.Selection.EndLine,
                        reference.Selection.EndColumn);

                    if (reference.QualifiedModuleName.Equals(selection.QualifiedName) &&
                        completeSelection.Contains(selection.Selection))
                    {
                        return declaration;
                    }
                }
            }

            return null;
        }

        public static IEnumerable<Declaration> FindInterfaces(this IEnumerable<Declaration> declarations)
        {
            var classes = declarations.Where(item => item.DeclarationType == DeclarationType.ClassModule);
            var interfaces = classes.Where(item => ((ClassModuleDeclaration)item).Subtypes.Any(s => s.IsUserDefined));
            return interfaces;
        }

        /// <summary>
        /// Finds all class members that are interface implementation members.
        /// </summary>
        public static IEnumerable<Declaration> FindInterfaceImplementationMembers(this IEnumerable<Declaration> declarations)
        {
            var items = declarations.ToList();
            var members = FindInterfaceMembers(items);
            var result = items.Where(item =>
                item.IsUserDefined
                && DeclarationExtensions.ProcedureTypes.Contains(item.DeclarationType)
                && members.Select(m => m.ComponentName + '_' + m.IdentifierName).Contains(item.IdentifierName))
            .ToList();

            return result;
        }

        //TODO: This looks incorrect.
        public static IEnumerable<Declaration> FindInterfaceImplementationMembers(this IEnumerable<Declaration> declarations, string interfaceMember)
        {
            return FindInterfaceImplementationMembers(declarations)
                .Where(m => m.IdentifierName.EndsWith(interfaceMember));
        }

        /// <summary>
        /// Locates all concrete implementations of the passed interface declaration.
        /// </summary>
        /// <param name="declarations">The declarations to search in.</param>
        /// <param name="interfaceDeclaration">The interface member to find.</param>
        /// <returns>All concrete implementations of the passed interface declaration.</returns>
        public static IEnumerable<Declaration> FindInterfaceImplementationMembers(this IEnumerable<Declaration> declarations, Declaration interfaceDeclaration)
        {
            return FindInterfaceImplementationMembers(declarations)
                .Where(decl => decl.ImplementsInterfaceMember(interfaceDeclaration));
        }

        //TODO: This looks incorrect.
        public static Declaration FindInterfaceMember(this IEnumerable<Declaration> declarations, Declaration implementation)
        {
            var members = FindInterfaceMembers(declarations);
            var matches = members.Where(m => m.IsUserDefined
                                             && m.DeclarationType == implementation.DeclarationType
                                             && implementation.IdentifierName == m.ComponentName + '_' + m.IdentifierName).ToList();

            return matches.Count > 1
                ? matches.SingleOrDefault(m => m.ProjectId == implementation.ProjectId)
                : matches.FirstOrDefault();
        }

        /// <summary>
        /// Finds all interface members.
        /// </summary>
        public static IEnumerable<Declaration> FindInterfaceMembers(this IEnumerable<Declaration> declarations)
        {
            var items = declarations.ToList();
            var interfaces = FindInterfaces(items).Select(i => i.Scope).ToList();
            var interfaceMembers = items.Where(item => item.IsUserDefined
                                                && (
                                                        DeclarationExtensions.ProcedureTypes.Contains(item.DeclarationType)
                                                        || item.DeclarationType == DeclarationType.Variable
                                                        && (item.Accessibility == Accessibility.Public || item.Accessibility == Accessibility.Implicit)
                                                   )
                                                && interfaces.Any(i => item.ParentScope.StartsWith(i)))
                                                .ToList();
            return interfaceMembers;
        }

        /// <summary>
        /// Finds all interface members defined by the passed decalaration.
        /// </summary>
        /// <param name="declarations">The declarations to search.</param>
        /// <param name="interfaceClass">The interface to find members from.</param>
        /// <returns>All interface members defined in interfaceClass</returns>
        public static IEnumerable<Declaration> FindInterfaceMembers(this IEnumerable<Declaration> declarations, Declaration interfaceClass)
        {
            var members = declarations.Where(decl => ReferenceEquals(decl.ParentDeclaration, interfaceClass) &&
                    (decl.Accessibility == Accessibility.Public || decl.Accessibility == Accessibility.Implicit) && 
                    (DeclarationExtensions.ProcedureTypes.Contains(decl.DeclarationType) || decl.DeclarationType == DeclarationType.Variable));
                
            return members;
        }
        /// <summary>
        /// Tests to see if a member Declaration is a concrete implemention of an interface member.
        /// </summary>
        /// <param name="member">The member to test for implementation.</param>
        /// <param name="interfaceMember">The interface member to test for implementation of.</param>
        /// <returns></returns>
        public static bool ImplementsInterfaceMember(this Declaration member, Declaration interfaceMember)
        {
            return (interfaceMember.Accessibility == Accessibility.Public || interfaceMember.Accessibility == Accessibility.Implicit)
                   && member.ParentDeclaration is ClassModuleDeclaration declaration
                   && declaration.Supertypes.Contains(interfaceMember.ParentDeclaration)
                   && DeclarationExtensions.ProcedureTypes.Contains(member.DeclarationType)
                   && (interfaceMember.DeclarationType == DeclarationType.Variable || DeclarationExtensions.ProcedureTypes.Contains(interfaceMember.DeclarationType))
                   && member.IdentifierName.Equals($"{interfaceMember.ComponentName}_{interfaceMember.IdentifierName}")
                   && (
                       member.DeclarationType == interfaceMember.DeclarationType
                       || (
                           interfaceMember.DeclarationType == DeclarationType.Variable
                           && (
                                member.DeclarationType == DeclarationType.PropertyGet
                                || (member.DeclarationType == DeclarationType.PropertyLet && !interfaceMember.IsObject)
                                || (member.DeclarationType == DeclarationType.PropertySet && (interfaceMember.IsObject || interfaceMember.AsTypeName.Equals(Tokens.Variant)))
                              )
                           )
                       );
        }
    }
}
