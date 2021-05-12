using System;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Parsing.VBA
{
    public class SelectedDeclarationProvider : ISelectedDeclarationProvider
    {
        private readonly ISelectionProvider _selectionProvider;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public SelectedDeclarationProvider(ISelectionProvider selectionProvider, IDeclarationFinderProvider finderProvider)
        {
            _selectionProvider = selectionProvider;
            _declarationFinderProvider = finderProvider;
        }

        public Declaration SelectedDeclaration()
        {
            return FromActiveSelection(SelectedDeclaration)();
        }

        private Func<T> FromActiveSelection<T>(Func<QualifiedSelection, T> func)
            where T: class
        {
            return () =>
            {
                var activeSelection = _selectionProvider.ActiveSelection();
                return activeSelection.HasValue
                    ? func(activeSelection.Value)
                    : null;
            };
        }

        public Declaration SelectedDeclaration(QualifiedModuleName module)
        {
            return FromModuleSelection(SelectedDeclaration)(module);
        }

        private Func<QualifiedModuleName, T> FromModuleSelection<T>(Func<QualifiedSelection, T> func) 
            where T : class
        {
            return (module) =>
            {
                var selection = _selectionProvider.Selection(module);
                if (!selection.HasValue)
                {
                    return null;
                }
                var qualifiedSelection = new QualifiedSelection(module, selection.Value);
                return func(qualifiedSelection);
            };
        }

        public Declaration SelectedDeclaration(QualifiedSelection qualifiedSelection)
        {
            var finder = _declarationFinderProvider?.DeclarationFinder;

            var candidateViaReference = SelectedDeclarationViaReference(qualifiedSelection, finder);
            if (candidateViaReference != null)
            {
                return candidateViaReference;
            }

            var candidateViaDeclaration = SelectedDeclarationViaDeclaration(qualifiedSelection, finder);
            if (candidateViaDeclaration != null)
            {
                return candidateViaDeclaration;
            }

            var candidateViaVariableDeclaration = SelectedDeclarationViaVariableDeclarationStatement(qualifiedSelection, finder);
            if (candidateViaVariableDeclaration != null)
            {
                return candidateViaVariableDeclaration;
            }

            var candidateViaConstantDeclaration = SelectedDeclarationViaConstantDeclarationStatement(qualifiedSelection, finder);
            if (candidateViaConstantDeclaration != null)
            {
                return candidateViaConstantDeclaration;
            }

            var candidateViaArgumentCallSite = SelectedDeclarationViaArgument(qualifiedSelection, finder);
            if (candidateViaArgumentCallSite != null)
            {
                return candidateViaArgumentCallSite;
            }

            // fallback to the containing member declaration if we're inside a procedure scope
            var candidateViaContainingMember = SelectedMember(qualifiedSelection);
            if (candidateViaContainingMember != null)
            {
                return candidateViaContainingMember;
            }

            // otherwise fallback to the containing module declaration
            return SelectedModule(qualifiedSelection);
        }

        private static Declaration SelectedDeclarationViaArgument(QualifiedSelection qualifiedSelection, DeclarationFinder finder)
        {
            var members = finder.Members(qualifiedSelection.QualifiedName)
                .Where(m => (m.DeclarationType.HasFlag(DeclarationType.Procedure) // includes PropertyLet and PropertySet and LibraryProcedure
                    || m.DeclarationType.HasFlag(DeclarationType.Function)) // includes PropertyGet and LibraryFunction
                    && !m.DeclarationType.HasFlag(DeclarationType.LibraryFunction)
                    && !m.DeclarationType.HasFlag(DeclarationType.LibraryProcedure));
            var enclosingProcedure = members.SingleOrDefault(m => m.Context.GetSelection().Contains(qualifiedSelection.Selection));
            if (enclosingProcedure == null)
            {
                return null;
            }

            var allArguments = enclosingProcedure.Context.GetDescendents<VBAParser.ArgumentContext>();

            var context = allArguments
                .Where(arg => arg.missingArgument() == null)
                .FirstOrDefault(m =>
                {
                    var isOnWhitespace = false;
                    if (m.TryGetPrecedingContext<VBAParser.WhiteSpaceContext>(out var whitespace))
                    {
                        isOnWhitespace = whitespace.GetSelection().ContainsFirstCharacter(qualifiedSelection.Selection);
                    }
                    return isOnWhitespace || m.GetSelection().ContainsFirstCharacter(qualifiedSelection.Selection);
                });
                
            var skippedArg = allArguments
                .Where(arg => arg.missingArgument() != null)
                .FirstOrDefault(m =>
                {
                    var isOnWhitespace = false;
                    if (m.TryGetPrecedingContext<VBAParser.WhiteSpaceContext>(out var whitespace))
                    {
                        isOnWhitespace = whitespace.GetSelection().ContainsFirstCharacter(qualifiedSelection.Selection);
                    }
                    return isOnWhitespace || m.GetSelection().ContainsFirstCharacter(qualifiedSelection.Selection);
                });

            context = context ?? skippedArg;
            if (context != null)
            {
                return (Declaration)finder.FindParameterOfNonDefaultMemberFromSimpleArgumentNotPassedByValExplicitly(context, enclosingProcedure)
                    ?? finder.FindInvokedMemberFromArgumentContext(context, qualifiedSelection.QualifiedName); // fallback to the invoked procedure declaration
            }

            return null;
        }

        private static Declaration SelectedDeclarationViaReference(QualifiedSelection qualifiedSelection, DeclarationFinder finder)
        {
            var referencesInModule = finder.IdentifierReferences(qualifiedSelection.QualifiedName);
            return referencesInModule
                .Where(reference => reference.IsSelected(qualifiedSelection))
                .Select(reference => reference.Declaration)
                .OrderByDescending(declaration => declaration.DeclarationType)
                // they're sorted by type, so a local comes before the procedure it's in
                .FirstOrDefault();
        }

        private static Declaration SelectedDeclarationViaDeclaration(QualifiedSelection qualifiedSelection, DeclarationFinder finder)
        {
            //There cannot be the identifier of a reference at this selection, but the module itself has this selection.
            //Resolving to the module would skip several valid alternatives.
            if (qualifiedSelection.Selection.Equals(Selection.Home))
            {
                return null;
            }

            var declarationsInModule = finder.Members(qualifiedSelection.QualifiedName);
            return declarationsInModule
                .Where(declaration => declaration.IsSelected(qualifiedSelection))
                .OrderByDescending(declaration => declaration.DeclarationType)
                // they're sorted by type, so a local comes before the procedure it's in
                .FirstOrDefault();
        }

        private static Declaration SelectedDeclarationViaVariableDeclarationStatement(QualifiedSelection qualifiedSelection, DeclarationFinder finder)
        {
            var variablesInModule = finder.Members(qualifiedSelection.QualifiedName)
                .Where(declaration => declaration.DeclarationType == DeclarationType.Variable);

            //This is annoying to do in method syntax LINQ. So this FirstOrDefault is done by hand.
            foreach (var variableDeclaration in variablesInModule)
            {
                var declarationSelection = SingleVariableDeclarationStatementSelection(variableDeclaration.Context as VBAParser.VariableSubStmtContext);
                if (declarationSelection.HasValue && declarationSelection.Value.Contains(qualifiedSelection.Selection))
                {
                    return variableDeclaration;
                }
            }

            return null;
        }

        private static Selection? SingleVariableDeclarationStatementSelection(VBAParser.VariableSubStmtContext context)
        {
            if (context is null)
            {
                return null;
            }

            var declaredVariableList = (VBAParser.VariableListStmtContext) context.Parent;
            if (declaredVariableList.variableSubStmt().Length != 1)
            {
                return null;
            }

            var declarationContext = (VBAParser.VariableStmtContext) declaredVariableList.Parent;
            return declarationContext.GetSelection();
        }

        private static Declaration SelectedDeclarationViaConstantDeclarationStatement(QualifiedSelection qualifiedSelection, DeclarationFinder finder)
        {
            var constantsInModule = finder.Members(qualifiedSelection.QualifiedName)
                .Where(declaration => declaration.DeclarationType == DeclarationType.Constant);

            //This is annoying to do in method syntax LINQ. So this FirstOrDefault is done by hand.
            foreach (var constantDeclaration in constantsInModule)
            {
                var declarationSelection = SingleConstantDeclarationStatementSelection(constantDeclaration.Context as VBAParser.ConstSubStmtContext);
                if (declarationSelection.HasValue && declarationSelection.Value.Contains(qualifiedSelection.Selection))
                {
                    return constantDeclaration;
                }
            }

            return null;
        }

        private static Selection? SingleConstantDeclarationStatementSelection(VBAParser.ConstSubStmtContext context)
        {
            if (context is null)
            {
                return null;
            }

            var declarationContext = (VBAParser.ConstStmtContext)context.Parent;
            if (declarationContext.constSubStmt().Length != 1)
            {
                return null;
            }

            return declarationContext.GetSelection();
        }

        public ModuleBodyElementDeclaration SelectedMember()
        {
            return FromActiveSelection(SelectedMember)();
        }

        public ModuleBodyElementDeclaration SelectedMember(QualifiedModuleName module)
        {
            return FromModuleSelection(SelectedMember)(module);
        }

        public ModuleBodyElementDeclaration SelectedMember(QualifiedSelection qualifiedSelection)
        {
            return _declarationFinderProvider.DeclarationFinder?
                .UserDeclarations(DeclarationType.Member)
                .OfType<ModuleBodyElementDeclaration>()
                .FirstOrDefault(member => member.QualifiedModuleName.Equals(qualifiedSelection.QualifiedName)
                                          && member.Context.GetSelection().Contains(qualifiedSelection.Selection));
        }

        public ModuleDeclaration SelectedModule()
        {
            return FromActiveSelection(SelectedModule)();
        }

        public ModuleDeclaration SelectedModule(QualifiedSelection qualifiedSelection)
        {
            return _declarationFinderProvider.DeclarationFinder?
                .UserDeclarations(DeclarationType.Module)
                .OfType<ModuleDeclaration>()
                .FirstOrDefault(module => module.QualifiedModuleName.Equals(qualifiedSelection.QualifiedName));
        }

        public ModuleDeclaration SelectedProjectExplorerModule()
        {
            var moduleName = _selectionProvider.ProjectExplorerSelection();
            return _declarationFinderProvider.DeclarationFinder?
                .ModuleDeclaration(moduleName) as ModuleDeclaration;
        }

        public ProjectDeclaration SelectedProject()
        {
            return FromActiveSelection(SelectedProject)();
        }

        public ProjectDeclaration SelectedProject(QualifiedSelection qualifiedSelection)
        {
            return _declarationFinderProvider.DeclarationFinder?
                .UserDeclarations(DeclarationType.Project)
                .OfType<ProjectDeclaration>()
                .FirstOrDefault(project => project.ProjectId.Equals(qualifiedSelection.QualifiedName.ProjectId));
        }
    }
}