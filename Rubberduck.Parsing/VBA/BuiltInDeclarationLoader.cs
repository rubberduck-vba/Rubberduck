using System;
using System.Collections.Generic;
using System.Linq;
using NLog;
using Rubberduck.Parsing.Symbols.DeclarationLoaders;
using System.Threading;

namespace Rubberduck.Parsing.VBA
{
    public class BuiltInDeclarationLoader : IBuiltInDeclarationLoader
    {
        private readonly IEnumerable<ICustomDeclarationLoader> _customDeclarationLoaders;
        private readonly RubberduckParserState _state;

        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public BuiltInDeclarationLoader(RubberduckParserState state, IEnumerable<ICustomDeclarationLoader> customDeclarationLoaders)
        {
            _state = state ?? throw new ArgumentNullException(nameof(state));

            _customDeclarationLoaders = customDeclarationLoaders ?? throw new ArgumentNullException(nameof(customDeclarationLoaders));
        }

        public bool LastLoadOfBuiltInDeclarationsLoadedDeclarations { get; private set; }

        public void LoadBuitInDeclarations()
        {
            LastLoadOfBuiltInDeclarationsLoadedDeclarations = false;
            foreach (var customDeclarationLoader in _customDeclarationLoaders)
            {
                try
                {
                    var customDeclarations = customDeclarationLoader.Load();
                    if (!customDeclarations.Any())
                    {
                        continue;
                    }

                    LastLoadOfBuiltInDeclarationsLoadedDeclarations = true;
                    foreach (var declaration in customDeclarations)
                    {
                        _state.AddDeclaration(declaration);
                    }
                }
                catch (Exception exception)
                {
                    Logger.Error(exception, "Exception thrown loading built-in declarations. (thread {0}).", Thread.CurrentThread.ManagedThreadId);
                }
            }
        }
    }
}
