using System;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Root
{
    public interface IInstanceProvider
    {
        RubberduckParserState StateInstance { get; }
    }

    public static class InstanceProviderFactory
    {
        public static IInstanceProvider GetInstanceProvider => new InstanceProvider();
    }

    internal class InstanceProvider : IInstanceProvider
    {
        private static RubberduckParserState _stateInstance;
        public static RubberduckParserState StateInstance
        {
            get => _stateInstance;
            set => _stateInstance = value ?? throw new NullReferenceException();
        }

        RubberduckParserState IInstanceProvider.StateInstance => _stateInstance;
    }
}
