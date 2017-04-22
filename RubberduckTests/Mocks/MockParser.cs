using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.Symbols.DeclarationLoaders;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Globalization;
using System.Reflection;
using System.Threading;
using Rubberduck.Parsing.PreProcessing;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.Mocks
{
    public static class MockParser
    {
        public static RubberduckParserState ParseString(string inputCode, out QualifiedModuleName qualifiedModuleName)
        {

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            qualifiedModuleName = new QualifiedModuleName(component);
            var parser = Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }
            return parser.State;

        }

        public static ParseCoordinator Create(IVBE vbe, RubberduckParserState state, string serializedDeclarationsPath = null)
        {
            var attributeParser = new Mock<IAttributeParser>();
            attributeParser.Setup(m => m.Parse(It.IsAny<IVBComponent>(), It.IsAny<CancellationToken>()))
                           .Returns(() => new Dictionary<Tuple<string, DeclarationType>, Attributes>());
            return Create(vbe, state, attributeParser.Object, serializedDeclarationsPath);
        }

        public static ParseCoordinator Create(IVBE vbe, RubberduckParserState state, IAttributeParser attributeParser, string serializedDeclarationsPath = null)
        {
            var path = serializedDeclarationsPath ??
                       Path.Combine(Path.GetDirectoryName(Assembly.GetAssembly(typeof(MockParser)).Location), "TestFiles", "Resolver");
            Func<IVBAPreprocessor> preprocessorFactory = () => new VBAPreprocessor(double.Parse(vbe.Version, CultureInfo.InvariantCulture));
            var projectManager = new SynchronousProjectManager(state, vbe);
            var moduleToModuleReferenceManager = new ModuleToModuleReferenceManager();
            var parserStateManager = new SynchronousParserStateManager(state);
            var referenceRemover = new SynchronousReferenceRemover(state, moduleToModuleReferenceManager);
            var comSynchronizer = new SynchronousCOMReferenceSynchronizer(
                state, 
                parserStateManager, 
                path);
            var builtInDeclarationLoader = new BuiltInDeclarationLoader(
                state,
                new List<ICustomDeclarationLoader>
                {
                    new DebugDeclarations(state),
                    new SpecialFormDeclarations(state),
                    new FormEventDeclarations(state),
                    new AliasDeclarations(state),
                });
            var parseRunner = new SynchronousParseRunner(
                state,
                parserStateManager,
                preprocessorFactory,
                attributeParser);
            var declarationResolveRunner = new SynchronousDeclarationResolveRunner(
                state, 
                parserStateManager, 
                comSynchronizer);
            var referenceResolveRunner = new SynchronousReferenceResolveRunner(
                state,
                parserStateManager,
                moduleToModuleReferenceManager,
                referenceRemover);
            var parsingStageService = new ParsingStageService(
                comSynchronizer,
                builtInDeclarationLoader,
                parseRunner,
                declarationResolveRunner,
                referenceResolveRunner
                );

            return new ParseCoordinator( 
                state,
                parsingStageService,
                projectManager,
                moduleToModuleReferenceManager,
                parserStateManager,
                referenceRemover,
                true);
        }

        public static RubberduckParserState CreateAndParse(IVBE vbe, string serializedDeclarationsPath = null)
        {
            var parser = Create(vbe, new RubberduckParserState(vbe));
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            return parser.State;
        }

        private static readonly HashSet<DeclarationType> ProceduralTypes =
            new HashSet<DeclarationType>(new[]
            {
                DeclarationType.Procedure, DeclarationType.Function, DeclarationType.PropertyGet,
                DeclarationType.PropertyLet, DeclarationType.PropertySet
            });

        public static void AddTestLibrary(this RubberduckParserState state, string serialized)
        {
            var reader = new XmlPersistableDeclarations();
            var deserialized = reader.Load(Path.Combine("Resolver", serialized));
            AddTestLibrary(state, deserialized);
        }

        // ReSharper disable once UnusedMember.Global; used by RubberduckWeb to load serialized declarations.
        public static void AddTestLibrary(this RubberduckParserState state, Stream stream)
        {
            var reader = new XmlPersistableDeclarations();
            var deserialized = reader.Load(stream);
            AddTestLibrary(state, deserialized);
        }

        private static void AddTestLibrary(RubberduckParserState state, SerializableProject deserialized)
        {
            var declarations = deserialized.Unwrap();

            foreach (var members in declarations.Where(d => d.DeclarationType != DeclarationType.Project &&
                                                            d.ParentDeclaration.DeclarationType == DeclarationType.ClassModule &&
                                                            ProceduralTypes.Contains(d.DeclarationType))
                .GroupBy(d => d.ParentDeclaration))
            {
                state.CoClasses.TryAdd(members.Select(m => m.IdentifierName).ToList(), members.First().ParentDeclaration);
            }

            foreach (var declaration in declarations)
            {
                state.AddDeclaration(declaration);
            }
        }
    }
}
