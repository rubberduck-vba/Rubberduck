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
        private RubberduckParserState _state;

        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public BuiltInDeclarationLoader(RubberduckParserState state, IEnumerable<ICustomDeclarationLoader> customDeclarationLoaders)
        {
            if (state == null) throw new ArgumentNullException(nameof(state));
            if (customDeclarationLoaders == null) throw new ArgumentNullException(nameof(customDeclarationLoaders));

            _state = state;
            _customDeclarationLoaders = customDeclarationLoaders;
        }

        private bool _lastExecutionLoadedDeclarations;
        public bool LastLoadOfBuiltInDeclarationsLoadedDeclarations
        {
            get
            {
                return _lastExecutionLoadedDeclarations;
            }
        }

        public void LoadBuitInDeclarations()
        {
            _lastExecutionLoadedDeclarations = false;
            foreach (var customDeclarationLoader in _customDeclarationLoaders)
            {
                try
                {
                    var customDeclarations = customDeclarationLoader.Load();
                    if (customDeclarations.Any())
                    {
                        _lastExecutionLoadedDeclarations = true;
                        foreach (var declaration in customDeclarationLoader.Load())
                        {
                            _state.AddDeclaration(declaration);
                        }
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
