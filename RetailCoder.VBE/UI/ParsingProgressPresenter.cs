using System;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;

namespace Rubberduck.UI
{
    public class ParsingProgressPresenter
    {
        private readonly IRubberduckParser _parser;
        private readonly ProgressDialog _view;

        public ParsingProgressPresenter(IRubberduckParser parser)
        {
            _view = new ProgressDialog();
            _parser = parser;
            _parser.ParseStarted += _parser_ParseStarted;
            _parser.ParseCompleted += _parser_ParseCompleted;
        }

        private void _parser_ParseCompleted(object sender, ParseCompletedEventArgs e)
        {
            _view.Close();
        }

        private void _parser_ParseStarted(object sender, ParseStartedEventArgs e)
        {
            _view.Show();
        }

        public VBProjectParseResult Parse(VBProject project)
        {
            return _parser.Parse(project, this);
        }
    }
}
