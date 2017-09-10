using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Antlr4.Runtime.Tree;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
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
        private enum DeclarationSite
        {
            NotApplicable = 0,
            ReferencedProject = 1,
            OtherComponent = 2,
            SameComponent = 3
        }

        private class OptionPrivateModuleListener : VBAParserBaseListener
        {
            public List<VBAParser.ModuleContext> OptionPrivateModules { get; } = new List<VBAParser.ModuleContext>();

            public override void EnterModule(VBAParser.ModuleContext context)
            {
                if (context.FindChildren<VBAParser.OptionPrivateModuleStmtContext>().Any())
                {
                    OptionPrivateModules.Add(context);
                }
            }
        }

        public ShadowedDeclarationInspection(RubberduckParserState state) : base(state)
        {
        }

        public override Type Type => typeof(ShadowedDeclarationInspection);

        public override CodeInspectionType InspectionType { get; } = CodeInspectionType.CodeQualityIssues;

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            var listener = new OptionPrivateModuleListener();
            var moduleDeclarations = State.DeclarationFinder.AllModules.Where(m => m.ComponentType == ComponentType.StandardModule);

            foreach (var module in moduleDeclarations)
            {
                ParseTreeWalker.Default.Walk(listener, State.GetParseTree(module));
            }

            var builtInEventHandlers = State.DeclarationFinder.FindEventHandlers().ToHashSet();

            var issues = new List<IInspectionResult>();

            var allUserProjects = UserDeclarations.OfType(DeclarationType.Project).Cast<ProjectDeclaration>();

            foreach (var userProject in allUserProjects)
            {
                var referencedProjectIds = userProject.ProjectReferences.Select(reference => reference.ReferencedProjectId).ToHashSet();

                var userDeclarations = UserDeclarations.Where(d =>
                    d.ProjectId == userProject.ProjectId &&
                    // User has no control over build-in event handlers or their parameters, so we skip them
                    !DeclarationIsPartOfBuiltInEventHandler(d, builtInEventHandlers));

                foreach (var declaration in userDeclarations)
                {
                    var shadowedDeclaration = State.AllDeclarations.FirstOrDefault(d =>
                        !Equals(d, declaration) &&
                        string.Equals(d.IdentifierName, declaration.IdentifierName, StringComparison.OrdinalIgnoreCase) &&
                        DeclarationCanBeShadowed(d, declaration, GetDeclarationSite(d, declaration, referencedProjectIds), listener));

                    if (shadowedDeclaration != null)
                    {
                        issues.Add(new DeclarationInspectionResult(this,
                            string.Format(InspectionsUI.ShadowedDeclarationInspectionResultFormat,
                                RubberduckUI.ResourceManager.GetString("DeclarationType_" + declaration.DeclarationType, CultureInfo.CurrentUICulture),
                                declaration.IdentifierName,
                                RubberduckUI.ResourceManager.GetString("DeclarationType_" + shadowedDeclaration.DeclarationType, CultureInfo.CurrentUICulture),
                                shadowedDeclaration.IdentifierName),
                            declaration));
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

            if (originalDeclaration.QualifiedName.QualifiedModuleName.Name != userDeclaration.QualifiedName.QualifiedModuleName.Name)
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

            var parameterDeclaration = declaration as ParameterDeclaration;

            return parameterDeclaration != null && builtInEventHandlers.Contains(parameterDeclaration.ParentDeclaration);
        }

        private static bool DeclarationCanBeShadowed(Declaration originalDeclaration, Declaration userDeclaration, DeclarationSite originalDeclarationSite, OptionPrivateModuleListener listener)
        {
            if (originalDeclarationSite == DeclarationSite.NotApplicable)
            {
                return false;
            }

            if (originalDeclarationSite == DeclarationSite.ReferencedProject)
            {
                return DeclarationInReferencedProjectCanBeShadowed(originalDeclaration, userDeclaration, listener);
            }

            if (originalDeclarationSite == DeclarationSite.OtherComponent)
            {
                return DeclarationInAnotherComponentCanBeShadowed(originalDeclaration, userDeclaration, listener);
            }

            return DeclarationInTheSameComponentCanBeShadowed(originalDeclaration, userDeclaration);
        }

        // TODO: Remove code duplication
        private static bool DeclarationInReferencedProjectCanBeShadowed(Declaration originalDeclaration, Declaration userDeclaration, OptionPrivateModuleListener listener)
        {
            if (DeclarationIsInsideOptionPrivateModule(originalDeclaration, listener))
            {
                return false;
            }

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
                if (userDeclarationComponentType == ComponentType.UserForm && !ReferencedProjectTypeShadowingRelations[originalDeclaration.DeclarationType].Contains(DeclarationType.UserForm))
                {
                    return false;
                }

                if (userDeclarationComponentType == ComponentType.Document && !ReferencedProjectTypeShadowingRelations[originalDeclaration.DeclarationType].Contains(DeclarationType.Document))
                {
                    return false;
                }
            }

            if (userDeclaration.DeclarationType != DeclarationType.ClassModule ||
                (userDeclarationComponentType != ComponentType.UserForm && userDeclarationComponentType != ComponentType.Document))
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

            return DeclarationAccessibilityCanBeShadowed(originalDeclaration);
        }

        private static bool DeclarationInAnotherComponentCanBeShadowed(Declaration originalDeclaration, Declaration userDeclaration, OptionPrivateModuleListener listener)
        {
            if (DeclarationIsInsideOptionPrivateModule(originalDeclaration, listener))
            {
                return false;
            }

            if (DeclarationIsProjectOrComponent(originalDeclaration) && DeclarationIsProjectOrComponent(userDeclaration))
            {
                return false;
            }

            var originalDeclarationComponentType = originalDeclaration.QualifiedName.QualifiedModuleName.ComponentType;

            // It is not possible to directly access a Parameter, UDT Member or Label declared in another component
            if (originalDeclaration.DeclarationType == DeclarationType.Parameter || originalDeclaration.DeclarationType == DeclarationType.UserDefinedTypeMember ||
                originalDeclaration.DeclarationType == DeclarationType.LineLabel)
            {
                return false;
            }

            // It is not possible to directly access any declarations placed inside a Class Module
            if (originalDeclaration.DeclarationType != DeclarationType.ClassModule && originalDeclarationComponentType == ComponentType.ClassModule)
            {
                return false;
            }

            if (originalDeclaration.DeclarationType == DeclarationType.ClassModule)
            {
                // Syntax of instantiating a new class makes it impossible to be shadowed
                if (originalDeclarationComponentType == ComponentType.ClassModule)
                {
                    return false;
                }

                if (originalDeclarationComponentType == ComponentType.UserForm && 
                    !OtherComponentTypeShadowingRelations[DeclarationType.UserForm].Contains(userDeclaration.DeclarationType))
                {
                    return false;
                }

                if (originalDeclarationComponentType == ComponentType.Document && 
                    !OtherComponentTypeShadowingRelations[DeclarationType.Document].Contains(userDeclaration.DeclarationType))
                {
                    return false;
                }
            }
            else if (!OtherComponentTypeShadowingRelations[originalDeclaration.DeclarationType].Contains(userDeclaration.DeclarationType))
            {
                return false;
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
            if (originalDeclaration.DeclarationType == DeclarationType.ProceduralModule || originalDeclaration.DeclarationType == DeclarationType.ClassModule)
            {
                return false;
            }

            // Syntax of instantiating a new UDT makes it impossible to be shadowed.
            if (originalDeclaration.DeclarationType == DeclarationType.UserDefinedType)
            {
                return false;
            }

            if (originalDeclaration.DeclarationType == DeclarationType.Parameter || originalDeclaration.DeclarationType == DeclarationType.UserDefinedTypeMember ||
                originalDeclaration.DeclarationType == DeclarationType.LineLabel)
            {
                return false;
            }

            // Events don't have a body, so their parameters can't be accessed
            if (userDeclaration.DeclarationType == DeclarationType.Parameter && userDeclaration.ParentDeclaration.DeclarationType == DeclarationType.Event)
            {
                return false;
            }

            // TODO: Distinguish between private and local
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

        private static bool DeclarationIsInsideOptionPrivateModule(Declaration declaration, OptionPrivateModuleListener listener)
        {
            if (declaration.QualifiedName.QualifiedModuleName.ComponentType != ComponentType.StandardModule)
            {
                return false;
            }

            var moduleDeclaration = declaration as ProceduralModuleDeclaration;
            if (moduleDeclaration != null)
            {
                return moduleDeclaration.IsPrivateModule;
            }

            return listener.OptionPrivateModules.Any(moduleContext => ParserRuleContextHelper.HasParent(declaration.Context, moduleContext));
        }

        private static bool DeclarationIsProjectOrComponent(Declaration declaration)
        {
            return declaration.DeclarationType == DeclarationType.Project ||
                   declaration.DeclarationType == DeclarationType.ProceduralModule ||
                   declaration.DeclarationType == DeclarationType.ClassModule ||
                   declaration.DeclarationType == DeclarationType.UserForm ||
                   declaration.DeclarationType == DeclarationType.Document;
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
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Document, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.ClassModule] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.ClassModule, DeclarationType.UserForm, DeclarationType.Document,
                }.ToHashSet(),
            [DeclarationType.Procedure] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Document, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.Function] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Document, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.PropertyGet] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Document, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.PropertySet] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Document, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.PropertyLet] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Document, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.Variable] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Document, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.Constant] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Document, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.Enumeration] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Document, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.EnumerationMember] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Document, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.UserDefinedType] = new[]
                {
                    DeclarationType.UserDefinedType
                }.ToHashSet(),
            [DeclarationType.LibraryProcedure] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Document, DeclarationType.Procedure, DeclarationType.Function,
                    DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter, DeclarationType.Variable, DeclarationType.Constant,
                    DeclarationType.Enumeration, DeclarationType.EnumerationMember, DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.LibraryFunction] = new[]
                {
                    DeclarationType.Project, DeclarationType.ProceduralModule, DeclarationType.UserForm, DeclarationType.Document, DeclarationType.Procedure, DeclarationType.Function,
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
                    DeclarationType.Parameter//, DeclarationType.Variable, DeclarationType.Constant
                }.ToHashSet(),
            [DeclarationType.Function] = new[]
                {
                    DeclarationType.Parameter//, DeclarationType.Variable, DeclarationType.Constant
                }.ToHashSet(),
            [DeclarationType.PropertyGet] = new[]
                {
                    DeclarationType.Parameter//, DeclarationType.Variable, DeclarationType.Constant
                }.ToHashSet(),
            [DeclarationType.PropertySet] = new[]
                {
                    DeclarationType.Parameter//, DeclarationType.Variable, DeclarationType.Constant
                }.ToHashSet(),
            [DeclarationType.PropertyLet] = new[]
                {
                    DeclarationType.Parameter//, DeclarationType.Variable, DeclarationType.Constant
                }.ToHashSet(),
            [DeclarationType.Variable] = new[]
                {
                    DeclarationType.Parameter, /*DeclarationType.Variable, DeclarationType.Constant,*/ DeclarationType.Enumeration
                }.ToHashSet(),
            [DeclarationType.Constant] = new[]
                {
                    DeclarationType.Parameter, /*DeclarationType.Variable, DeclarationType.Constant,*/ DeclarationType.Enumeration
                }.ToHashSet(),
            [DeclarationType.Enumeration] = new[]
                {
                    DeclarationType.Procedure, DeclarationType.Function, DeclarationType.PropertyGet, DeclarationType.PropertySet, DeclarationType.PropertyLet, DeclarationType.Parameter,
                    /*DeclarationType.Variable, DeclarationType.Constant,*/ DeclarationType.LibraryProcedure, DeclarationType.LibraryFunction
                }.ToHashSet(),
            [DeclarationType.EnumerationMember] = new[]
                {
                    DeclarationType.Parameter//, DeclarationType.Variable, DeclarationType.Constant
                }.ToHashSet(),
            [DeclarationType.LibraryProcedure] = new[]
                {
                    DeclarationType.Parameter//, DeclarationType.Variable, DeclarationType.Constant
                }.ToHashSet(),
            [DeclarationType.LibraryFunction] = new[]
                {
                    DeclarationType.Parameter//, DeclarationType.Variable, DeclarationType.Constant
                }.ToHashSet()
        };
    }
}
