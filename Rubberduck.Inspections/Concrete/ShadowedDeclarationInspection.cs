using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class ShadowedDeclarationInspection : InspectionBase
    {
        public ShadowedDeclarationInspection(RubberduckParserState state) : base(state)
        {
        }

        public override CodeInspectionType InspectionType { get; } = CodeInspectionType.CodeQualityIssues;

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            var builtInEventHandlers = State.DeclarationFinder.FindEventHandlers().ToHashSet();

            var issues = new List<IInspectionResult>();

            var allUserProjects = UserDeclarations.OfType(DeclarationType.Project).Cast<ProjectDeclaration>();

            foreach (var userProject in allUserProjects)
            {
                var referencedProjectIds = userProject.ProjectReferences.Select(reference => reference.ReferencedProjectId).ToHashSet();

                var userDeclarations = UserDeclarations.Where(d => d.ProjectId == userProject.ProjectId);

                // User has no control over build-in event handlers or their parameters, so we skip them
                userDeclarations = userDeclarations.Where(d => !DeclarationIsPartOfBuiltInEventHandler(d, builtInEventHandlers));

                foreach (var declaration in userDeclarations)
                {
                    var shadowedDeclaration = State.AllDeclarations.FirstOrDefault(d =>
                        referencedProjectIds.Contains(d.ProjectId) && d.IdentifierName == declaration.IdentifierName && DeclarationCanBeShadowed(d, declaration));

                    if (shadowedDeclaration != null)
                    {
                        issues.Add(new DeclarationInspectionResult(this,
                            string.Format(InspectionsUI.ShadowedDeclarationInspectionResultFormat,
                                RubberduckUI.ResourceManager.GetString("DeclarationType_" + declaration.DeclarationType, CultureInfo.CurrentUICulture),
                                declaration.IdentifierName,
                                RubberduckUI.ResourceManager.GetString("DeclarationType_" + shadowedDeclaration.DeclarationType,CultureInfo.CurrentUICulture),
                                shadowedDeclaration.IdentifierName),
                            declaration));
                    }
                } 
            }

            return issues;
        }

        private static bool DeclarationIsPartOfBuiltInEventHandler(Declaration declaration, ICollection<Declaration> builtInEventHandlers)
        {
            if (builtInEventHandlers.Contains(declaration))
            {
                return true;
            }

            var parameterDeclaration = declaration as ParameterDeclaration;

            return parameterDeclaration != null && builtInEventHandlers.Contains(parameterDeclaration.ParentDeclaration);
        }

        private bool DeclarationCanBeShadowed(Declaration originalDeclaration, Declaration userDeclaration)
        {
            var originalDeclarationComponentType = originalDeclaration.QualifiedName.QualifiedModuleName.ComponentType;
            var userDeclarationComponentType = userDeclaration.QualifiedName.QualifiedModuleName.ComponentType;

            // It is not possible to directly access a Parameter, UDT Member or Label declared in another project
            if (originalDeclaration.DeclarationType == DeclarationType.Parameter || originalDeclaration.DeclarationType == DeclarationType.UserDefinedTypeMember || 
                originalDeclaration.DeclarationType == DeclarationType.LineLabel)
            {
                return false;
            }

            // It is not possible to instantiate a Class Module which is not exposed
            if ((originalDeclaration as ClassModuleDeclaration)?.IsExposed == false)
            {
               return false;
            }

            // It is not possible to directly access a UserForm or Document declared in another project, nor any declarations placed inside them
            if (originalDeclarationComponentType == ComponentType.UserForm || originalDeclarationComponentType == ComponentType.Document)
            {
                return false;
            }

            // It is not possible to directly access any declarations placed inside a Class Module
            if (originalDeclaration.DeclarationType != DeclarationType.ClassModule && originalDeclarationComponentType == ComponentType.ClassModule)
            {
                return false;
            }

            if (userDeclaration.DeclarationType == DeclarationType.ClassModule)
            {
                if (userDeclarationComponentType == ComponentType.UserForm && !TypeShadowingRelations[originalDeclaration.DeclarationType].Contains(DeclarationType.UserForm))
                {
                    return false;
                }

                if (userDeclarationComponentType == ComponentType.Document && !TypeShadowingRelations[originalDeclaration.DeclarationType].Contains(DeclarationType.Document))
                {
                    return false;
                }
            }

            if (userDeclaration.DeclarationType != DeclarationType.ClassModule || (userDeclarationComponentType != ComponentType.UserForm && userDeclarationComponentType != ComponentType.Document))
            {
                if (!TypeShadowingRelations[originalDeclaration.DeclarationType].Contains(userDeclaration.DeclarationType))
                {
                    return false;
                }
            }

            // Events don't have a body, so their parameters can't be accessed
            if (userDeclaration.DeclarationType == DeclarationType.Parameter && userDeclaration.ParentDeclaration.DeclarationType == DeclarationType.Event)
            {
                return false;
            }

            return originalDeclaration.Accessibility == Accessibility.Global ||
                   originalDeclaration.Accessibility == Accessibility.Public ||
                   originalDeclaration.DeclarationType == DeclarationType.Project ||
                   // Enumeration member can be shadowed only when enclosing enumeration has public accessibility
                   (originalDeclaration.DeclarationType == DeclarationType.EnumerationMember && originalDeclaration.ParentDeclaration.Accessibility == Accessibility.Public);
        }

        // Dictionary values represents all declaration types that can shadow the declaration type of the key
        private static readonly Dictionary<DeclarationType, HashSet<DeclarationType>> TypeShadowingRelations = new Dictionary
            <DeclarationType, HashSet<DeclarationType>>
            {
                [DeclarationType.Project] = new[]
                {
                    DeclarationType.Procedure, DeclarationType.Function, DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet,
                    DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure,
                    DeclarationType.LibraryFunction, DeclarationType.LineLabel
                }.ToHashSet(),
                [DeclarationType.ProceduralModule] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Document, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction, DeclarationType.LineLabel
                }.ToHashSet(),
                [DeclarationType.ClassModule] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.ClassModule, DeclarationType.UserForm, DeclarationType.Document,
                }.ToHashSet(),
                [DeclarationType.Procedure] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Document, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction, DeclarationType.LineLabel
                }.ToHashSet(),
                [DeclarationType.Function] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Document, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction, DeclarationType.LineLabel
                }.ToHashSet(),
                [DeclarationType.PropertyGet] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Document, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction, DeclarationType.LineLabel
                }.ToHashSet(),
                [DeclarationType.PropertySet] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Document, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction, DeclarationType.LineLabel
                }.ToHashSet(),
                [DeclarationType.PropertyLet] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Document, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction, DeclarationType.LineLabel
                }.ToHashSet(),
                [DeclarationType.Variable] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Document, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction, DeclarationType.LineLabel
                }.ToHashSet(),
                [DeclarationType.Constant] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Document, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction, DeclarationType.LineLabel
                }.ToHashSet(),
                [DeclarationType.Enumeration] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Document, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction, DeclarationType.LineLabel
                }.ToHashSet(),
                [DeclarationType.EnumerationMember] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Document, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction, DeclarationType.LineLabel
                }.ToHashSet(),
                [DeclarationType.UserDefinedType] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Document, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction, DeclarationType.LineLabel
                }.ToHashSet(),
                [DeclarationType.LibraryProcedure] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Document, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction, DeclarationType.LineLabel
                }.ToHashSet(),
                [DeclarationType.LibraryFunction] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Document, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction, DeclarationType.LineLabel
                }.ToHashSet(),
            };
    }
}
