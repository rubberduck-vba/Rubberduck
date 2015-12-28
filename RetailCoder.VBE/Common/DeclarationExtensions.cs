using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Windows.Media.Imaging;
using Microsoft.Vbe.Interop;
using Rubberduck.Annotations;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
// ReSharper disable LocalizableElement

namespace Rubberduck.Common
{
    public static class DeclarationExtensions
    {
        private static readonly DeclarationIconCache Cache = new DeclarationIconCache();

        public static BitmapImage BitmapImage(this Declaration declaration)
        {
            return Cache[declaration];
        }

        /// <summary>
        /// Returns the Selection of a VariableStmtContext.
        /// </summary>
        /// <exception cref="ArgumentException">Throws when target's DeclarationType is not Variable.</exception>
        /// <param name="target"></param>
        /// <returns></returns>
        public static Selection GetVariableStmtContextSelection(this Declaration target)
        {
            if (target.DeclarationType != DeclarationType.Variable)
            {
                throw new ArgumentException("Target DeclarationType is not Variable.", "target");
            }

            var statement = GetVariableStmtContext(target);

            return new Selection(statement.Start.Line, statement.Start.Column,
                    statement.Stop.Line, statement.Stop.Column);
        }

        /// <summary>
        /// Returns a VariableStmtContext.
        /// </summary>
        /// <exception cref="ArgumentException">Throws when target's DeclarationType is not Variable.</exception>
        /// <param name="target"></param>
        /// <returns></returns>
        public static VBAParser.VariableStmtContext GetVariableStmtContext(this Declaration target)
        {
            if (target.DeclarationType != DeclarationType.Variable)
            {
                throw new ArgumentException("Target DeclarationType is not Variable.", "target");
            }

            var statement = target.Context.Parent.Parent as VBAParser.VariableStmtContext;
            if (statement == null)
            {
                throw new MissingMemberException("Statement not found");
            }

            return statement;
        }

        /// <summary>
        /// Returns whether a variable declaration statement contains multiple declarations in a single statement.
        /// </summary>
        /// <exception cref="ArgumentException">Throws when target's DeclarationType is not Variable.</exception>
        /// <param name="target"></param>
        /// <returns></returns>
        public static bool HasMultipleDeclarationsInStatement(this Declaration target)
        {
            if (target.DeclarationType != DeclarationType.Variable)
            {
                throw new ArgumentException("Target DeclarationType is not Variable.", "target");
            }

            var statement = target.Context.Parent as VBAParser.VariableListStmtContext;

            return statement != null && statement.children.OfType<VBAParser.VariableSubStmtContext>().Count() > 1;
        }

        /// <summary>
        /// Returns the number of variable declarations in a single statement.
        /// </summary>
        /// <exception cref="ArgumentException">Throws when target's DeclarationType is not Variable.</exception>
        /// <param name="target"></param>
        /// <returns></returns>
        public static int CountOfDeclarationsInStatement(this Declaration target)
        {
            if (target.DeclarationType != DeclarationType.Variable)
            {
                throw new ArgumentException("Target DeclarationType is not Variable.", "target");
            }

            var statement = target.Context.Parent as VBAParser.VariableListStmtContext;

            if (statement != null)
            {
                return statement.children.OfType<VBAParser.VariableSubStmtContext>().Count();
            }

            throw new ArgumentException("'target.Context.Parent' is not type VBAParser.VariabelListStmtContext", "target");
        }

        /// <summary>
        /// Returns the number of variable declarations in a single statement.  Adjusted to be 1-indexed rather than 0-indexed.
        /// </summary>
        /// <exception cref="ArgumentException">Throws when target's DeclarationType is not Variable.</exception>
        /// <param name="target"></param>
        /// <returns></returns>
        public static int IndexOfVariableDeclarationInStatement(this Declaration target)
        {
            if (target.DeclarationType != DeclarationType.Variable)
            {
                throw new ArgumentException("Target DeclarationType is not Variable.", "target");
            }

            var statement = target.Context.Parent as VBAParser.VariableListStmtContext;

            if (statement != null)
            {
                return statement.children.OfType<VBAParser.VariableSubStmtContext>()
                        .ToList()
                        .IndexOf((VBAParser.VariableSubStmtContext)target.Context) + 1;
            }

            // ReSharper disable once LocalizableElement
            throw new ArgumentException("'target.Context.Parent' is not type VBAParser.VariabelListStmtContext", "target");
        }

        public static readonly DeclarationType[] ProcedureTypes =
        {
            DeclarationType.Procedure,
            DeclarationType.Function,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet
        };

        /// <summary>
        /// Gets all declarations of the specified <see cref="DeclarationType"/>.
        /// </summary>
        public static IEnumerable<Declaration> OfType(this IEnumerable<Declaration> declarations, DeclarationType declarationType)
        {
            return declarations.Where(declaration =>
                declaration.DeclarationType == declarationType);
        }

        /// <summary>
        /// Gets all declarations of any one of the specified <see cref="DeclarationType"/> values.
        /// </summary>
        public static IEnumerable<Declaration> OfType(this IEnumerable<Declaration> declarations, params DeclarationType[] declarationTypes)
        {
            return declarations.Where(declaration =>
                declarationTypes.Any(type => declaration.DeclarationType == type));
        }

        public static IEnumerable<Declaration> Named(this IEnumerable<Declaration> declarations, string name)
        {
            return declarations.Where(declaration => declaration.IdentifierName == name);
        }

        /// <summary>
        /// Gets the declaration for all identifiers declared in or below the specified scope.
        /// </summary>
        public static IEnumerable<Declaration> InScope(this IEnumerable<Declaration> declarations, string scope)
        {
            return string.IsNullOrEmpty(scope) 
                ? declarations 
                : declarations.Where(declaration => declaration.Scope.StartsWith(scope));
        }

        /// <summary>
        /// Gets the declaration for all identifiers declared in or below the specified scope.
        /// </summary>
        public static IEnumerable<Declaration> InScope(this IEnumerable<Declaration> declarations, [NotNull] Declaration parent)
        {
            return declarations.Where(declaration => declaration.ParentScope == parent.Scope);
        }

        public static IEnumerable<Declaration> FindInterfaces(this IEnumerable<Declaration> declarations)
        {
            var classes = declarations.Where(item => item.DeclarationType == DeclarationType.Class);
            var interfaces = classes.Where(item => item.References.Any(reference =>
                reference.Context.Parent is VBAParser.ImplementsStmtContext));

            return interfaces;
        }

        /// <summary>
        /// Finds all interface members.
        /// </summary>
        public static IEnumerable<Declaration> FindInterfaceMembers(this IEnumerable<Declaration> declarations)
        {
            var items = declarations.ToList();
            var interfaces = FindInterfaces(items).Select(i => i.Scope).ToList();
            var interfaceMembers = items.Where(item => !item.IsBuiltIn
                                                && ProcedureTypes.Contains(item.DeclarationType)
                                                && interfaces.Any(i => item.ParentScope.StartsWith(i)))
                                                .ToList();
            return interfaceMembers;
        }

        /// <summary>
        /// Finds all event handler procedures for specified control declaration.
        /// </summary>
        public static IEnumerable<Declaration> FindEventHandlers(this IEnumerable<Declaration> declarations, Declaration control)
        {
            Debug.Assert(control.DeclarationType == DeclarationType.Control);

            return declarations.Where(declaration => declaration.ParentScope == control.ParentScope
                && declaration.DeclarationType == DeclarationType.Procedure
                && declaration.IdentifierName.StartsWith(control.IdentifierName + "_"));
        }

        /// <summary>
        /// Gets the <see cref="Declaration"/> of the specified <see cref="type"/>, 
        /// at the specified <see cref="selection"/>.
        /// Returns the declaration if selection is on an identifier reference.
        /// </summary>
        public static Declaration FindSelectedDeclaration(this IEnumerable<Declaration> declarations, QualifiedSelection selection, DeclarationType type, Func<Declaration, Selection> selector = null)
        {
            return FindSelectedDeclaration(declarations, selection, new[] { type }, selector);
        }

        /// <summary>
        /// Gets the <see cref="Declaration"/> of the specified <see cref="types"/>, 
        /// at the specified <see cref="selection"/>.
        /// Returns the declaration if selection is on an identifier reference.
        /// </summary>
        public static Declaration FindSelectedDeclaration(this IEnumerable<Declaration> declarations, QualifiedSelection selection, IEnumerable<DeclarationType> types, Func<Declaration, Selection> selector = null)
        {
            var userDeclarations = declarations.Where(item => !item.IsBuiltIn);
            var items = userDeclarations.Where(item => types.Contains(item.DeclarationType)
                && item.QualifiedName.QualifiedModuleName == selection.QualifiedName).ToList();

            var declaration = items.SingleOrDefault(item =>
                selector == null
                    ? item.Selection.Contains(selection.Selection)
                    : selector(item).Contains(selection.Selection));

            if (declaration != null)
            {
                return declaration;
            }

            // if we haven't returned yet, then we must be on an identifier reference.
            declaration = items.SingleOrDefault(item => !item.IsBuiltIn
                && types.Contains(item.DeclarationType)
                && item.References.Any(reference =>
                reference.QualifiedModuleName == selection.QualifiedName
                && reference.Selection.Contains(selection.Selection)));

            return declaration;
        }

        public static IEnumerable<Declaration> FindFormEventHandlers(this IEnumerable<Declaration> declarations)
        {
            var items = declarations.ToList();

            var forms = items.Where(item => item.DeclarationType == DeclarationType.Class
                && item.QualifiedName.QualifiedModuleName.Component != null
                && item.QualifiedName.QualifiedModuleName.Component.Type == vbext_ComponentType.vbext_ct_MSForm)
                .ToList();

            var result = new List<Declaration>();
            foreach (var declaration in forms)
            {
                result.AddRange(FindFormEventHandlers(items, declaration));
            }

            return result;
        }

        public static IEnumerable<Declaration> FindFormEventHandlers(this IEnumerable<Declaration> declarations, Declaration userForm)
        {
            var items = declarations as IList<Declaration> ?? declarations.ToList();
            var events = items.Where(item => item.IsBuiltIn
                                                     && item.ParentScope == "MSForms.UserForm"
                                                     && item.DeclarationType == DeclarationType.Event).ToList();
            var handlerNames = events.Select(item => "UserForm_" + item.IdentifierName);
            var handlers = items.Where(item => item.ParentScope == userForm.Scope
                                                       && item.DeclarationType == DeclarationType.Procedure
                                                       && handlerNames.Contains(item.IdentifierName));

            return handlers.ToList();
        }

            /// <summary>
        /// Gets a tuple containing the <c>WithEvents</c> declaration and the corresponding handler,
        /// for each type implementing this event.
        /// </summary>
        public static IEnumerable<Tuple<Declaration,Declaration>> FindHandlersForEvent(this IEnumerable<Declaration> declarations, Declaration eventDeclaration)
        {
            var items = declarations as IList<Declaration> ?? declarations.ToList();
            return items.Where(item => item.IsWithEvents && item.AsTypeName == eventDeclaration.ComponentName)
            .Select(item => new
            {
                WithEventDeclaration = item, 
                EventProvider = items.SingleOrDefault(type => type.DeclarationType == DeclarationType.Class && type.QualifiedName.QualifiedModuleName == item.QualifiedName.QualifiedModuleName)
            })
            .Select(item => new
            {
                WithEventsDeclaration = item.WithEventDeclaration,
                ProviderEvents = items.Where(member => member.DeclarationType == DeclarationType.Event && member.QualifiedSelection.QualifiedName == item.EventProvider.QualifiedName.QualifiedModuleName)
            })
            .Select(item => Tuple.Create(
                item.WithEventsDeclaration,
                items.SingleOrDefault(declaration => declaration.DeclarationType == DeclarationType.Procedure
                && declaration.QualifiedName.QualifiedModuleName == item.WithEventsDeclaration.QualifiedName.QualifiedModuleName
                && declaration.IdentifierName == item.WithEventsDeclaration.IdentifierName + '_' + eventDeclaration.IdentifierName)
                ));
        }

        public static IEnumerable<Declaration> FindEventProcedures(this IEnumerable<Declaration> declarations, Declaration withEventsDeclaration)
        {
            if (!withEventsDeclaration.IsWithEvents)
            {
                return new Declaration[]{};
            }

            var items = declarations as IList<Declaration> ?? declarations.ToList();
            var type = items.SingleOrDefault(item => item.DeclarationType == DeclarationType.Class
                                                             && item.Project != null
                                                             && item.IdentifierName == withEventsDeclaration.AsTypeName.Split('.').Last());

            if (type == null)
            {
                return new Declaration[]{};
            }

            var members = GetTypeMembers(items, type).ToList();
            var events = members.Where(member => member.DeclarationType == DeclarationType.Event);
            var handlerNames = events.Select(e => withEventsDeclaration.IdentifierName + '_' + e.IdentifierName);

            return items.Where(item => item.Project != null 
                                               && item.Project.Equals(withEventsDeclaration.Project)
                                               && item.ParentScope == withEventsDeclaration.ParentScope
                                               && item.DeclarationType == DeclarationType.Procedure
                                               && handlerNames.Any(name => item.IdentifierName == name))
                .ToList();
        }

        private static IEnumerable<Declaration> GetTypeMembers(this IEnumerable<Declaration> declarations, Declaration type)
        {
            return declarations.Where(item => item.Project != null && item.Project.Equals(type.Project) && item.ParentScope == type.Scope);
        }

            /// <summary>
        /// Finds all class members that are interface implementation members.
        /// </summary>
        public static IEnumerable<Declaration> FindInterfaceImplementationMembers(this IEnumerable<Declaration> declarations)
        {
            var items = declarations.ToList();
            var members = FindInterfaceMembers(items);
            var result = items.Where(item => 
                !item.IsBuiltIn 
                && ProcedureTypes.Contains(item.DeclarationType)
                && members.Select(m => m.ComponentName + '_' + m.IdentifierName).Contains(item.IdentifierName))
            .ToList();

            return result;
        }

        public static IEnumerable<Declaration> FindInterfaceImplementationMembers(this IEnumerable<Declaration> declarations, string interfaceMember)
        {
            return FindInterfaceImplementationMembers(declarations)
                .Where(m => m.IdentifierName.EndsWith(interfaceMember));
        }

        public static Declaration FindInterfaceMember(this IEnumerable<Declaration> declarations, Declaration implementation)
        {
            var members = FindInterfaceMembers(declarations);
            var matches = members.Where(m => !m.IsBuiltIn && implementation.IdentifierName == m.ComponentName + '_' + m.IdentifierName).ToList();

            return matches.Count > 1 
                ? matches.SingleOrDefault(m => m.Project == implementation.Project) 
                : matches.First();
        }

        /// <summary>
        /// Returns the declaration contained in a qualified selection.
        /// To get the selection of a variable or field, use FindVariable(QualifiedSelection)
        /// </summary>
        /// <param name="declarations"></param>
        /// <param name="selection"></param>
        /// <param name="validDeclarationTypes"></param>
        /// <returns></returns>
        public static Declaration FindTarget(this IEnumerable<Declaration> declarations, QualifiedSelection selection, DeclarationType[] validDeclarationTypes)
        {
            var items = declarations.ToList();

            var target = items
                .Where(item => !item.IsBuiltIn)
                .FirstOrDefault(item => item.IsSelected(selection)
                                     || item.References.Any(r => r.IsSelected(selection)));

            if (target != null && validDeclarationTypes.Contains(target.DeclarationType))
            {
                return target;
            }

            target = null;

            var targets = items
                .Where(item => !item.IsBuiltIn
                               && item.ComponentName == selection.QualifiedName.ComponentName
                               && validDeclarationTypes.Contains(item.DeclarationType));

            var currentSelection = new Selection(0, 0, int.MaxValue, int.MaxValue);

            foreach (var declaration in targets)
            {
                var activeSelection = new Selection(declaration.Context.Start.Line,
                                                    declaration.Context.Start.Column,
                                                    declaration.Context.Stop.Line,
                                                    declaration.Context.Stop.Column);

                if (currentSelection.Contains(activeSelection) && activeSelection.Contains(selection.Selection))
                {
                    target = declaration;
                    currentSelection = activeSelection;
                }

                foreach (var reference in declaration.References)
                {
                    var proc = (dynamic)reference.Context.Parent;
                    VBAParser.ArgsCallContext paramList;

                    // This is to prevent throws when this statement fails:
                    // (VBAParser.ArgsCallContext)proc.argsCall();
                    try { paramList = (VBAParser.ArgsCallContext)proc.argsCall(); }
                    catch { continue; }

                    if (paramList == null) { continue; }

                    activeSelection = new Selection(paramList.Start.Line,
                                                    paramList.Start.Column,
                                                    paramList.Stop.Line,
                                                    paramList.Stop.Column + paramList.Stop.Text.Length + 1);

                    if (currentSelection.Contains(activeSelection) && activeSelection.Contains(selection.Selection))
                    {
                        target = reference.Declaration;
                        currentSelection = activeSelection;
                    }
                }
            }
            return target;
        }

        /// <summary>
        /// Returns the variable which contains the passed-in QualifiedSelection.  Returns null if the selection is not on a variable.
        /// </summary>
        /// <param name="declarations"></param>
        /// <param name="selection"></param>
        /// <returns></returns>
        public static Declaration FindVariable(this IEnumerable<Declaration> declarations, QualifiedSelection selection)
        {
            var items = declarations.Where(d => !d.IsBuiltIn && d.DeclarationType == DeclarationType.Variable).ToList();

            var target = items
                .FirstOrDefault(item => item.IsSelected(selection) || item.References.Any(r => r.IsSelected(selection)));

            if (target != null) { return target; }

            var targets = items.Where(item => item.ComponentName == selection.QualifiedName.ComponentName);

            foreach (var declaration in targets)
            {
                var declarationSelection = new Selection(declaration.Context.Start.Line,
                                                    declaration.Context.Start.Column,
                                                    declaration.Context.Stop.Line,
                                                    declaration.Context.Stop.Column + declaration.Context.Stop.Text.Length);

                if (declarationSelection.Contains(selection.Selection) ||
                    !HasMultipleDeclarationsInStatement(declaration) && GetVariableStmtContextSelection(declaration).Contains(selection.Selection))
                {
                    return declaration;
                }

                var reference =
                    declaration.References.FirstOrDefault(r => r.Selection.Contains(selection.Selection));

                if (reference != null)
                {
                    return reference.Declaration;
                }
            }
            return null;
        }

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
                    var implementsStmt = reference.Context.Parent as VBAParser.ImplementsStmtContext;

                    if (implementsStmt == null) { continue; }

                    var completeSelection = new Selection(implementsStmt.GetSelection().StartLine,
                        implementsStmt.GetSelection().StartColumn, reference.Selection.EndLine,
                        reference.Selection.EndColumn);

                    if (reference.QualifiedModuleName.ComponentName == selection.QualifiedName.ComponentName &&
                        reference.QualifiedModuleName.Project == selection.QualifiedName.Project &&
                        completeSelection.Contains(selection.Selection))
                    {
                        return declaration;
                    }
                }
            }

            return null;
        }
    }
}
