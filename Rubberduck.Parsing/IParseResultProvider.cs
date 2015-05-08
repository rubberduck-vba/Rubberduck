using System;
using System.Collections.Generic;

namespace Rubberduck.Parsing
{
    public class ParseStartedEventArgs : EventArgs
    {
        public ParseStartedEventArgs(IEnumerable<string> projectNames)
        {
            _projectNames = projectNames;
        }

        private readonly IEnumerable<string> _projectNames;
        public IEnumerable<string> ProjectNames { get { return _projectNames; } }
    }

    public class ParseCompletedEventArgs : EventArgs
    {
        public ParseCompletedEventArgs(IEnumerable<VBProjectParseResult> results)
        {
            _results = results;
        }

        private readonly IEnumerable<VBProjectParseResult> _results;
        public IEnumerable<VBProjectParseResult> ParseResults { get { return _results; } }
    }

    public interface IParseResultProvider
    {
        event EventHandler<ParseStartedEventArgs> ParseStarted;
        event EventHandler<ParseCompletedEventArgs> ParseCompleted;
    }
}
