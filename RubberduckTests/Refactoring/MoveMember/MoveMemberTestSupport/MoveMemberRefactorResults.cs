using Rubberduck.Refactorings.MoveMember;
using System.Collections.Generic;

namespace RubberduckTests.Refactoring.MoveMember
{
    public struct MoveMemberRefactorResults
    {
        private readonly IDictionary<string, string> _results;
        private readonly string _sourceModuleName;
        private readonly string _destinationModuleName;
        private readonly string _strategyName;

        public MoveMemberRefactorResults(MoveEndpoints endpoints, IDictionary<string, string> refactorResults)
            :this(endpoints.SourceModuleName(), endpoints.DestinationModuleName(), refactorResults)
        {}

        public MoveMemberRefactorResults(string sourceModuleName, string destinationModuleName, IDictionary<string, string> refactorResults)
        {
            _results = refactorResults;
            _sourceModuleName = sourceModuleName;
            _destinationModuleName = destinationModuleName;
            _strategyName = nameof(MoveMemberToStdModule);
        }

        public string this[string moduleName]
        {
            get => _results[moduleName];
        }

        public string Source => _results[_sourceModuleName];
        public string Destination => _results[_destinationModuleName];
        public string StrategyName => _strategyName;
    }
}
