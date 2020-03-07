using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;

namespace Rubberduck.CodeAnalysis.Inspections.Abstract
{
    internal abstract class IsMissingInspectionBase : ArgumentReferenceInspectionFromDeclarationsBase<ParameterDeclaration>
    {
        protected IsMissingInspectionBase(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {}

        private static readonly List<string> IsMissingQualifiedNames = new List<string>
        {
            "VBE7.DLL;VBA.Information.IsMissing",
            "VBA6.DLL;VBA.Information.IsMissing"
        };

        protected override IEnumerable<Declaration> ObjectionableDeclarations(DeclarationFinder finder)
        {
            return IsMissingDeclarations(finder);
        }

        protected IReadOnlyList<Declaration> IsMissingDeclarations(DeclarationFinder finder)
        {
            var vbaProjects = finder.Projects
                .Where(project => project.IdentifierName == "VBA" && !project.IsUserDefined)
                .ToList();

            if (!vbaProjects.Any())
            {
                return new List<Declaration>();
            }

            var informationModules = vbaProjects
                .Select(project => finder.FindStdModule("Information", project, true))
                .OfType<ModuleDeclaration>()
                .ToList();

            if (!informationModules.Any())
            {
                return new List<Declaration>();
            }

            var isMissing = informationModules
                .SelectMany(module => module.Members)
                .Where(decl => IsMissingQualifiedNames.Contains(decl.QualifiedName.ToString()))
                .OfType<ModuleBodyElementDeclaration>();

            return isMissing
                .SelectMany(declaration => declaration.Parameters)
                .ToList();
        }

        protected ParameterDeclaration ParameterForReference(ArgumentReference reference, DeclarationFinder finder)
        {
            var argumentContext = reference.Context as VBAParser.LExprContext;
            if (!(argumentContext?.lExpression() is VBAParser.SimpleNameExprContext name))
            {
                return null;
            }

            var procedure = reference.Context.GetAncestor<VBAParser.ModuleBodyElementContext>();
            var module = reference.QualifiedModuleName;

            return finder.Members(module, DeclarationType.Member)
                .OfType<ModuleBodyElementDeclaration>()
                .FirstOrDefault(decl => decl.Context.Parent == procedure)?.Parameters
                    .FirstOrDefault(param => param.IdentifierName.Equals(name.GetText()));
        }
    }
}
