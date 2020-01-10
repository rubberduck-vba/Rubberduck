using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using NLog;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Binding;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;

namespace Rubberduck.Inspections.Inspections.Abstract
{
    public abstract class IsMissingInspectionBase : IdentifierReferenceInspectionFromDeclarationsBase
    {
        protected IsMissingInspectionBase(RubberduckParserState state) 
            : base(state) { }

        private static readonly List<string> IsMissingQualifiedNames = new List<string>
        {
            "VBE7.DLL;VBA.Information.IsMissing",
            "VBA6.DLL;VBA.Information.IsMissing"
        };

        protected abstract bool IsUnsuitableArgument(ArgumentReference reference, DeclarationFinder finder);


        protected override IEnumerable<Declaration> ObjectionableDeclarations(DeclarationFinder finder)
        {
            return IsMissingDeclarations(finder);
        }

        protected IReadOnlyList<Declaration> IsMissingDeclarations(DeclarationFinder finder)
        {
            var vbaProjects = finder.Projects
                .Where(project => project.IdentifierName == "VBA")
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
                .ToList();

            return isMissing;
        }

        protected override IEnumerable<IdentifierReference> ObjectionableReferences(DeclarationFinder finder)
        {
            return ObjectionableDeclarations(finder)
                .OfType<ModuleBodyElementDeclaration>()
                .SelectMany(declaration => declaration.Parameters)
                .SelectMany(parameter => parameter.ArgumentReferences)
                .Where(reference => IsResultReference(reference, finder));
        }

        protected override bool IsResultReference(IdentifierReference reference, DeclarationFinder finder)
        {
            return reference is ArgumentReference argumentReference
                   && IsUnsuitableArgument(argumentReference, finder);
        }

        protected ParameterDeclaration GetParameterForReference(ArgumentReference reference, DeclarationFinder finder)
        {
            var argumentContext = reference.Context as VBAParser.LExprContext;
            if (!(argumentContext?.lExpression() is VBAParser.SimpleNameExprContext name))
            {
                return null;
            }

            var procedure = reference.Context.GetAncestor<VBAParser.ModuleBodyElementContext>();
            //TODO: revisit this once PR #5338 is merged.
            return finder.UserDeclarations(DeclarationType.Member)
                .OfType<ModuleBodyElementDeclaration>()
                .FirstOrDefault(decl => decl.Context.Parent == procedure)?.Parameters
                    .FirstOrDefault(param => param.IdentifierName.Equals(name.GetText()));
        }
    }
}
