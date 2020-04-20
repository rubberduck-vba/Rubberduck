using Rubberduck.Refactorings.MoveMember;
using System.Collections.Generic;

namespace RubberduckTests.Refactoring.MoveMember
{
    public struct MoveMemberRefactorTestResults
    {
        private readonly IDictionary<string, string> _results;
        private readonly string _sourceModuleName;
        private readonly string _destinationModuleName;

        public MoveMemberRefactorTestResults(MoveEndpoints endpoints, IDictionary<string, string> refactorResults)
            :this(endpoints.SourceModuleName(), endpoints.DestinationModuleName(), refactorResults)
        {}

        public MoveMemberRefactorTestResults(string sourceModuleName, string destinationModuleName, IDictionary<string, string> refactorResults)
        {
            _results = refactorResults;
            _sourceModuleName = sourceModuleName;
            _destinationModuleName = destinationModuleName;
        }

        public string this[string moduleName]
        {
            get => _results[moduleName];
        }

        public string Source => _results[_sourceModuleName];
        public string Destination => _results[_destinationModuleName];
    }
}
