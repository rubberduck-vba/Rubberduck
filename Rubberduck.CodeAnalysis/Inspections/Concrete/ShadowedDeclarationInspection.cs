﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.Resources;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Identifies identifiers that hide/"shadow" other identifiers otherwise accessible in that scope.
    /// </summary>
    /// <why>
    /// Global namespace contains a number of perfectly legal identifier names that user code can use. But using these names in user code 
    /// effectively "hides" the global ones. In general, avoid shadowing global-scope identifiers if possible.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Private MsgBox As String ' hides the global-scope VBA.Interaction.MsgBox function in this module.
    /// 
    /// Public Sub DoSomething()
    ///     MsgBox = "Test" ' refers to the module variable in scope.
    ///     VBA.Interaction.MsgBox MsgBox ' global function now needs to be fully qualified to be accessed.
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Private message As String
    /// 
    /// Public Sub DoSomething()
    ///     message = "Test"
    ///     MsgBox message ' VBA.Interaction module qualifier is optional.
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class ShadowedDeclarationInspection : InspectionBase
    {
        private enum DeclarationSite
        {
            NotApplicable = 0,
            ReferencedProject = 1,
            OtherComponent = 2,
            SameComponent = 3
        }

        public ShadowedDeclarationInspection(RubberduckParserState state) : base(state)
        {
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var builtInEventHandlers = State.DeclarationFinder.FindEventHandlers().ToHashSet();

            var issues = new List<IInspectionResult>();

            var allUserProjects = State.DeclarationFinder.UserDeclarations(DeclarationType.Project).Cast<ProjectDeclaration>();

            foreach (var userProject in allUserProjects)
            {
                var referencedProjectIds = userProject.ProjectReferences.Select(reference => reference.ReferencedProjectId).ToHashSet();

                var userDeclarations = UserDeclarations.Where(declaration =>
                    declaration.ProjectId == userProject.ProjectId &&
                    // User has no control over build-in event handlers or their parameters, so we skip them
                    !DeclarationIsPartOfBuiltInEventHandler(declaration, builtInEventHandlers));

                foreach (var userDeclaration in userDeclarations)
                {
                    var shadowedDeclaration = State.DeclarationFinder
                        .MatchName(userDeclaration.IdentifierName).FirstOrDefault(declaration => 
                            !declaration.Equals(userDeclaration) &&
                            DeclarationCanBeShadowed(declaration, userDeclaration, GetDeclarationSite(declaration, userDeclaration, referencedProjectIds)));

                    if (shadowedDeclaration != null)
                    {
                        issues.Add(new DeclarationInspectionResult(this,
                            string.Format(Resources.Inspections.InspectionResults.ShadowedDeclarationInspection,
                                RubberduckUI.ResourceManager.GetString("DeclarationType_" + userDeclaration.DeclarationType, CultureInfo.CurrentUICulture),
                                userDeclaration.QualifiedName,
                                RubberduckUI.ResourceManager.GetString("DeclarationType_" + shadowedDeclaration.DeclarationType, CultureInfo.CurrentUICulture),
                                shadowedDeclaration.QualifiedName),
                            userDeclaration));
                    }
                }
            }

            return issues;
        }

        private static DeclarationSite GetDeclarationSite(Declaration originalDeclaration, Declaration userDeclaration, ICollection<string> referencedProjectIds)
        {
            if (originalDeclaration.ProjectId != userDeclaration.ProjectId)
            {
                return referencedProjectIds.Contains(originalDeclaration.ProjectId) ? DeclarationSite.ReferencedProject : DeclarationSite.NotApplicable;
            }

            if (originalDeclaration.QualifiedName.QualifiedModuleName.ComponentName != userDeclaration.QualifiedName.QualifiedModuleName.ComponentName)
            {
                return DeclarationSite.OtherComponent;
            }

            return DeclarationSite.SameComponent;
        }

        private static bool DeclarationIsPartOfBuiltInEventHandler(Declaration declaration, ICollection<Declaration> builtInEventHandlers)
        {
            if (builtInEventHandlers.Contains(declaration))
            {
                return true;
            }

            return declaration is ParameterDeclaration parameterDeclaration &&
                   builtInEventHandlers.Contains(parameterDeclaration.ParentDeclaration);
        }

        private static bool DeclarationCanBeShadowed(Declaration originalDeclaration, Declaration userDeclaration, DeclarationSite originalDeclarationSite)
        {
            switch (originalDeclarationSite)
            {
                case DeclarationSite.NotApplicable:
                    return false;
                case DeclarationSite.ReferencedProject:
                    return DeclarationInReferencedProjectCanBeShadowed(originalDeclaration, userDeclaration);
                case DeclarationSite.OtherComponent:
                    return DeclarationInAnotherComponentCanBeShadowed(originalDeclaration, userDeclaration);
            }

            return DeclarationInTheSameComponentCanBeShadowed(originalDeclaration, userDeclaration);
        }

        private static bool DeclarationInReferencedProjectCanBeShadowed(Declaration originalDeclaration, Declaration userDeclaration)
        {
            if (DeclarationIsInsideOptionPrivateModule(originalDeclaration))
            {
                return false;
            }

            if ((originalDeclaration.ParentDeclaration as ClassModuleDeclaration)?.IsGlobalClassModule == false)
            {
                return false;
            }

            var originalDeclarationEnclosingType = originalDeclaration.QualifiedName.QualifiedModuleName.ComponentType;
            var userDeclarationEnclosingType = userDeclaration.QualifiedName.QualifiedModuleName.ComponentType;

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
            if (originalDeclarationEnclosingType == ComponentType.UserForm || originalDeclarationEnclosingType == ComponentType.Document)
            {
                return false;
            }

            // It is not possible to directly access any declarations placed inside a Class Module
            if (originalDeclaration.DeclarationType != DeclarationType.ClassModule && originalDeclarationEnclosingType == ComponentType.ClassModule)
            {
                return false;
            }

            if (userDeclaration.DeclarationType == DeclarationType.ClassModule || userDeclaration.DeclarationType == DeclarationType.Document)
            {
                switch (userDeclarationEnclosingType)
                {
                    case ComponentType.UserForm when !ReferencedProjectTypeShadowingRelations[originalDeclaration.DeclarationType].Contains(DeclarationType.UserForm):
                        return false;
                    case ComponentType.Document when !ReferencedProjectTypeShadowingRelations[originalDeclaration.DeclarationType].Contains(DeclarationType.Document):
                        return false;
                }
            }

            if ((userDeclaration.DeclarationType != DeclarationType.ClassModule && userDeclaration.DeclarationType != DeclarationType.Document) ||
                (userDeclarationEnclosingType != ComponentType.UserForm && userDeclarationEnclosingType != ComponentType.Document))
            {
                if (!ReferencedProjectTypeShadowingRelations[originalDeclaration.DeclarationType].Contains(userDeclaration.DeclarationType))
                {
                    return false;
                }
            }

            // Events don't have a body, so their parameters can't be accessed
            if (userDeclaration.DeclarationType == DeclarationType.Parameter && userDeclaration.ParentDeclaration.DeclarationType == DeclarationType.Event)
            {
                return false;
            }

            // We insert Debug.Assert as a member access on an artificial Debug standard module. Thus, Assert will also be seen as shadowing Debug.Assert, which is not true. 
            if (originalDeclaration.IdentifierName.Equals("Assert", StringComparison.InvariantCultureIgnoreCase) && originalDeclaration.QualifiedModuleName.ComponentName.Equals("Debug", StringComparison.InvariantCultureIgnoreCase))
            {
                return false;
            }

            return DeclarationAccessibilityCanBeShadowed(originalDeclaration);
        }

        private static bool DeclarationInAnotherComponentCanBeShadowed(Declaration originalDeclaration, Declaration userDeclaration)
        {
            if (DeclarationIsInsideOptionPrivateModule(originalDeclaration))
            {
                return false;
            }

            if (DeclarationIsProjectOrComponent(originalDeclaration) && DeclarationIsProjectOrComponent(userDeclaration))
            {
                return false;
            }

            var originalDeclarationEnclosingType = originalDeclaration.QualifiedName.QualifiedModuleName.ComponentType;

            // It is not possible to directly access a Parameter, UDT Member or Label declared in another component.
            if (originalDeclaration.DeclarationType == DeclarationType.Parameter || originalDeclaration.DeclarationType == DeclarationType.UserDefinedTypeMember ||
                originalDeclaration.DeclarationType == DeclarationType.LineLabel)
            {
                return false;
            }

            // It is not possible to directly access any declarations placed inside a Class Module.
            if (originalDeclaration.DeclarationType != DeclarationType.ClassModule &&
                originalDeclaration.DeclarationType != DeclarationType.Document &&
                originalDeclarationEnclosingType == ComponentType.ClassModule)
            {
                return false;
            }

            // It is not possible to directly access any declarations placed inside a Document Module. (Document Modules have DeclarationType ClassMoodule.)
            if (originalDeclaration.DeclarationType != DeclarationType.ClassModule &&
                originalDeclaration.DeclarationType != DeclarationType.Document && 
                originalDeclarationEnclosingType == ComponentType.Document)
            {
                return false;
            }

            // It is not possible to directly access any declarations placed inside a User Form. (User Forms have DeclarationType ClassMoodule.)
            if (originalDeclaration.DeclarationType != DeclarationType.ClassModule &&
                originalDeclaration.DeclarationType != DeclarationType.Document && 
                originalDeclarationEnclosingType == ComponentType.UserForm)
            {
                return false;
            }

            if (originalDeclaration.DeclarationType == DeclarationType.ClassModule || originalDeclaration.DeclarationType == DeclarationType.Document)
            {
                // Syntax of instantiating a new class makes it impossible to be shadowed
                switch (originalDeclarationEnclosingType)
                {
                    case ComponentType.ClassModule:
                        return false;
                    case ComponentType.UserForm when !OtherComponentTypeShadowingRelations[DeclarationType.UserForm].Contains(userDeclaration.DeclarationType):
                        return false;
                    case ComponentType.Document when !OtherComponentTypeShadowingRelations[DeclarationType.Document].Contains(userDeclaration.DeclarationType):
                        return false;
                }
            }
            else
            {
                if (!OtherComponentTypeShadowingRelations.TryGetValue(originalDeclaration.DeclarationType,
                        out var shadowedTypes)
                    || !shadowedTypes.Contains(userDeclaration.DeclarationType))
                {
                    return false;
                }
            }

            // Events don't have a body, so their parameters can't be accessed
            if (userDeclaration.DeclarationType == DeclarationType.Parameter && userDeclaration.ParentDeclaration.DeclarationType == DeclarationType.Event)
            {
                return false;
            }

            return DeclarationAccessibilityCanBeShadowed(originalDeclaration);
        }

        private static bool DeclarationInTheSameComponentCanBeShadowed(Declaration originalDeclaration, Declaration userDeclaration)
        {
            // Shadowing the component containing the declaration is not a problem, because it is possible to directly access declarations inside that component
            if (originalDeclaration.DeclarationType == DeclarationType.ProceduralModule || 
                originalDeclaration.DeclarationType == DeclarationType.ClassModule ||
                originalDeclaration.DeclarationType == DeclarationType.Document ||
                userDeclaration.DeclarationType == DeclarationType.ProceduralModule || 
                userDeclaration.DeclarationType == DeclarationType.ClassModule ||
                userDeclaration.DeclarationType == DeclarationType.Document)
            {
                return false;
            }

            // Syntax of instantiating a new UDT makes it impossible to be shadowed
            switch (originalDeclaration.DeclarationType)
            {
                case DeclarationType.UserDefinedType:
                case DeclarationType.Parameter:
                case DeclarationType.UserDefinedTypeMember:
                case DeclarationType.LineLabel:
                    return false;
            }

            if ((originalDeclaration.DeclarationType == DeclarationType.Variable || originalDeclaration.DeclarationType == DeclarationType.Constant) &&
                DeclarationIsLocal(originalDeclaration))
            {
                return false;
            }

            if (userDeclaration.DeclarationType == DeclarationType.Variable || userDeclaration.DeclarationType == DeclarationType.Constant)
            {
                return DeclarationIsLocal(userDeclaration);
            }
            
            // Shadowing between two enumerations or enumeration members is not possible inside one component.
            if (((originalDeclaration.DeclarationType == DeclarationType.Enumeration 
                    && userDeclaration.DeclarationType == DeclarationType.EnumerationMember)
                || (originalDeclaration.DeclarationType == DeclarationType.EnumerationMember
                    && userDeclaration.DeclarationType == DeclarationType.Enumeration)))
            { 
                    var originalDeclarationIndex = originalDeclaration.Context.start.StartIndex;
                    var userDeclarationIndex = userDeclaration.Context.start.StartIndex;

                    // First declaration wins
                    return originalDeclarationIndex > userDeclarationIndex 
                           // Enumeration member can have the same name as enclosing enumeration
                           && !userDeclaration.Equals(originalDeclaration.ParentDeclaration);
            }

            // Events don't have a body, so their parameters can't be accessed
            if (userDeclaration.DeclarationType == DeclarationType.Parameter && userDeclaration.ParentDeclaration.DeclarationType == DeclarationType.Event)
            {
                return false;
            }

            return  SameComponentTypeShadowingRelations[originalDeclaration.DeclarationType].Contains(userDeclaration.DeclarationType);
        }

        private static bool DeclarationAccessibilityCanBeShadowed(Declaration originalDeclaration)
        {
            return originalDeclaration.Accessibility == Accessibility.Global ||
                   originalDeclaration.Accessibility == Accessibility.Public ||
                   originalDeclaration.DeclarationType == DeclarationType.Project ||
                   // Enumeration member can be shadowed only when enclosing enumeration has public accessibility
                   (originalDeclaration.DeclarationType == DeclarationType.EnumerationMember && originalDeclaration.ParentDeclaration.Accessibility == Accessibility.Public);
        }

        private static bool DeclarationIsInsideOptionPrivateModule(Declaration declaration)
        {
            if (declaration.QualifiedName.QualifiedModuleName.ComponentType != ComponentType.StandardModule)
            {
                return false;
            }

            if (Declaration.GetModuleParent(declaration) is ProceduralModuleDeclaration moduleDeclaration)
            {
                return moduleDeclaration.IsPrivateModule;
            }

            return false;
        }

        private static bool DeclarationIsProjectOrComponent(Declaration declaration)
        {
            return declaration.DeclarationType == DeclarationType.Project ||
                   declaration.DeclarationType == DeclarationType.ProceduralModule ||
                   declaration.DeclarationType == DeclarationType.ClassModule ||
                   declaration.DeclarationType == DeclarationType.UserForm ||
                   declaration.DeclarationType == DeclarationType.Document;
        }

        private static bool DeclarationIsLocal(Declaration declaration)
        {
            return declaration.ParentDeclaration.DeclarationType == DeclarationType.Procedure ||
                   declaration.ParentDeclaration.DeclarationType == DeclarationType.Function ||
                   declaration.ParentDeclaration.DeclarationType == DeclarationType.PropertyGet ||
                   declaration.ParentDeclaration.DeclarationType == DeclarationType.PropertySet ||
                   declaration.ParentDeclaration.DeclarationType == DeclarationType.PropertyLet;
        }

        // Dictionary values represent all declaration types that can shadow the declaration type of the key
        private static readonly Dictionary<DeclarationType, HashSet<DeclarationType>> ReferencedProjectTypeShadowingRelations = new Dictionary<DeclarationType, HashSet<DeclarationType>>
        {
            [DeclarationType.Project] = new[]
                {
                    DeclarationType.Procedure, DeclarationType.Function, DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet,
                    DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure,
                    DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.ProceduralModule] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.ClassModule] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.ClassModule, DeclarationType.UserForm, DeclarationType.Document
                }.ToHashSet(),
            [DeclarationType.Procedure] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.Function] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.PropertyGet] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.PropertySet] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.PropertyLet] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.Variable] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.Constant] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.Enumeration] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.EnumerationMember] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.UserDefinedType] = new[]
                {
                    DeclarationType.UserDefinedType
                }.ToHashSet(),
            [DeclarationType.LibraryProcedure] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.LibraryFunction] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet()
        };

        // Dictionary values represent all declaration types that can shadow the declaration type of the key
        private static readonly Dictionary<DeclarationType, HashSet<DeclarationType>> OtherComponentTypeShadowingRelations = new Dictionary<DeclarationType, HashSet<DeclarationType>>
        {
            [DeclarationType.Project] = new[]
                {
                    DeclarationType.Procedure, DeclarationType.Function, DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter,
                    DeclarationType.Variable, DeclarationType.Constant, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.ProceduralModule] = new[]
                {
                    DeclarationType.Procedure, DeclarationType.Function, DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter,
                    DeclarationType.Variable, DeclarationType.Constant, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.UserForm] = new[]
                {
                    DeclarationType.Procedure, DeclarationType.Function, DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter,
                    DeclarationType.Variable, DeclarationType.Constant, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.Document] = new[]
                {
                    DeclarationType.Procedure, DeclarationType.Function, DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter,
                    DeclarationType.Variable, DeclarationType.Constant, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.Procedure] = new[]
                {
                    DeclarationType.Procedure, DeclarationType.Function, DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter,
                    DeclarationType.Variable, DeclarationType.Constant, DeclarationType.Enumeration, DeclarationType.EnumerationMember,
                    DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.Function] = new[]
                {
                    DeclarationType.Procedure, DeclarationType.Function, DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter,
                    DeclarationType.Variable, DeclarationType.Constant, DeclarationType.Enumeration, DeclarationType.EnumerationMember,
                    DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.PropertyGet] = new[]
                {
                    DeclarationType.Procedure, DeclarationType.Function, DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter,
                    DeclarationType.Variable, DeclarationType.Constant, DeclarationType.Enumeration, DeclarationType.EnumerationMember,
                    DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.PropertySet] = new[]
                {
                    DeclarationType.Procedure, DeclarationType.Function, DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter,
                    DeclarationType.Variable, DeclarationType.Constant, DeclarationType.Enumeration, DeclarationType.EnumerationMember,
                    DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.PropertyLet] = new[]
                {
                    DeclarationType.Procedure, DeclarationType.Function, DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter,
                    DeclarationType.Variable, DeclarationType.Constant, DeclarationType.Enumeration, DeclarationType.EnumerationMember,
                    DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.Variable] = new[]
                {
                    DeclarationType.Procedure, DeclarationType.Function, DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter,
                    DeclarationType.Variable, DeclarationType.Constant, DeclarationType.Enumeration, DeclarationType.EnumerationMember,
                    DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.Constant] = new[]
                {
                    DeclarationType.Procedure, DeclarationType.Function, DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter,
                    DeclarationType.Variable, DeclarationType.Constant, DeclarationType.Enumeration, DeclarationType.EnumerationMember,
                    DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.Enumeration] = new[]
                {
                    DeclarationType.Procedure, DeclarationType.Function, DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter,
                    DeclarationType.Variable, DeclarationType.Constant, DeclarationType.Enumeration, DeclarationType.EnumerationMember,
                    DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.EnumerationMember] = new[]
                {
                    DeclarationType.Procedure, DeclarationType.Function, DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter,
                    DeclarationType.Variable, DeclarationType.Constant, DeclarationType.Enumeration, DeclarationType.EnumerationMember,
                    DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.UserDefinedType] = new[]
                {
                    DeclarationType.UserDefinedType
                }.ToHashSet(),
            [DeclarationType.LibraryProcedure] = new[]
                {
                    DeclarationType.Procedure, DeclarationType.Function, DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter,
                    DeclarationType.Variable, DeclarationType.Constant, DeclarationType.Enumeration, DeclarationType.EnumerationMember,
                    DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.LibraryFunction] = new[]
                {
                    DeclarationType.Procedure, DeclarationType.Function, DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter,
                    DeclarationType.Variable, DeclarationType.Constant, DeclarationType.Enumeration, DeclarationType.EnumerationMember,
                    DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet()
        };

        // Dictionary values represent all declaration types that can shadow the declaration type of the key
        private static readonly Dictionary<DeclarationType, HashSet<DeclarationType>> SameComponentTypeShadowingRelations = new Dictionary<DeclarationType, HashSet<DeclarationType>>
        {
            [DeclarationType.Procedure] = new[]
                {
                    DeclarationType.Parameter
                }.ToHashSet(),
            [DeclarationType.Function] = new[]
                {
                    DeclarationType.Parameter
                }.ToHashSet(),
            [DeclarationType.PropertyGet] = new[]
                {
                    DeclarationType.Parameter
                }.ToHashSet(),
            [DeclarationType.PropertySet] = new[]
                {
                    DeclarationType.Parameter
                }.ToHashSet(),
            [DeclarationType.PropertyLet] = new[]
                {
                    DeclarationType.Parameter
                }.ToHashSet(),
            [DeclarationType.Variable] = new[]
                {
                    DeclarationType.Parameter, DeclarationType.Enumeration
                }.ToHashSet(),
            [DeclarationType.Constant] = new[]
                {
                    DeclarationType.Parameter, DeclarationType.Enumeration
                }.ToHashSet(),
            [DeclarationType.Enumeration] = new[]
                {
                    DeclarationType.Procedure, DeclarationType.Function, DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter,
                    DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.EnumerationMember] = new[]
                {
                    DeclarationType.Parameter
                }.ToHashSet(),
            [DeclarationType.LibraryProcedure] = new[]
                {
                    DeclarationType.Parameter
                }.ToHashSet(),
            [DeclarationType.LibraryFunction] = new[]
                {
                    DeclarationType.Parameter
                }.ToHashSet()
        };
    }
}
