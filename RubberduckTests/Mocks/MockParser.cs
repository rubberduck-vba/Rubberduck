using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NUnit.Framework;
using Moq;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.Symbols.DeclarationLoaders;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Globalization;
using System.Reflection;
using System.Threading;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.PreProcessing;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.Parsing.Symbols.ParsingExceptions;

namespace RubberduckTests.Mocks
{
    public static class MockParser
    {
        public static RubberduckParserState ParseString(string inputCode, out QualifiedModuleName qualifiedModuleName)
        {

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            qualifiedModuleName = new QualifiedModuleName(component);
            var parser = Create(vbe.Object);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status == ParserState.Error)
            {
                Assert.Inconclusive("Parser Error: {0}");
            }
            return parser.State;

        }

        public static ParseCoordinator Create(IVBE vbe, string serializedDeclarationsPath = null)
        {
            var declarationFinderFactory = new DeclarationFinderFactory();
            var state = new RubberduckParserState(vbe, declarationFinderFactory);
            return Create(vbe, state, serializedDeclarationsPath);
        }

        public static ParseCoordinator Create(IVBE vbe, RubberduckParserState state, string serializedDeclarationsPath = null)
        {
            var attributeParser = new TestAttributeParser(() => new VBAPreprocessor(double.Parse(vbe.Version, CultureInfo.InvariantCulture)));
            var exporter = new Mock<IModuleExporter>().Object;
            return Create(vbe, state, attributeParser, exporter, serializedDeclarationsPath);
        }

        public static ParseCoordinator Create(IVBE vbe, RubberduckParserState state, IAttributeParser attributeParser, IModuleExporter exporter, string serializedDeclarationsPath = null)
        {
            var path = serializedDeclarationsPath ??
                       Path.Combine(Path.GetDirectoryName(Assembly.GetAssembly(typeof(MockParser)).Location), "TestFiles", "Resolver");
            Func<IVBAPreprocessor> preprocessorFactory = () => new VBAPreprocessor(double.Parse(vbe.Version, CultureInfo.InvariantCulture));
            var projectManager = new SynchronousProjectManager(state, vbe);
            var moduleToModuleReferenceManager = new ModuleToModuleReferenceManager();
            var supertypeClearer = new SynchronousSupertypeClearer(state); 
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
                attributeParser,
                exporter);
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
            var parsingCacheService = new ParsingCacheService(
                state,
                moduleToModuleReferenceManager,
                referenceRemover,
                supertypeClearer
                );

            return new ParseCoordinator(
                state,
                parsingStageService,
                parsingCacheService,
                projectManager,
                parserStateManager,
                true);
        }

        public static RubberduckParserState CreateAndParse(IVBE vbe, string serializedDeclarationsPath = null)
        {
            var parser = Create(vbe);
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
            var deserialized = reader.Load(Path.Combine("Testfiles//Resolver", serialized));
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

        public static RubberduckParserState CreateAndParse(IVBE vbe, IInspectionListener listener)
        {
            var parser = Create(vbe);
            parser.Parse(new CancellationTokenSource());
            if(parser.State.Status >= ParserState.Error)
            { Assert.Inconclusive("Parser Error"); }

            return parser.State;
        }
    }
}
