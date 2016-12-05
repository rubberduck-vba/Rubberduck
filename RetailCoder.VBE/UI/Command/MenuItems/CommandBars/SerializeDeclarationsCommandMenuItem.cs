using System;
using System.IO;
using System.Linq;
using NLog;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.UI.Command.MenuItems.CommandBars
{
    public class SerializeDeclarationsCommandMenuItem : CommandMenuItemBase
    {
        public SerializeDeclarationsCommandMenuItem(CommandBase command) : base(command)
        {
        }

        public override Func<string> Caption { get { return () => "Serialize"; } }
        public override string Key { get { return "SerializeDeclarations"; } }
    }

    public class SerializeDeclarationsCommand : CommandBase
    {
        private readonly RubberduckParserState _state;

        public SerializeDeclarationsCommand(RubberduckParserState state) 
            : base(LogManager.GetCurrentClassLogger())
        {
            _state = state;
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            return _state.Status == ParserState.Ready;
        }

        protected override void ExecuteImpl(object parameter)
        {
            var declarations = _state.AllDeclarations
                .Where(declaration => declaration.IsBuiltIn)
                .Select(declaration => new SerializableDeclaration(declaration))
                .GroupBy(declaration => declaration.QualifiedMemberName.QualifiedModuleName.ProjectPath);
            foreach (var project in declarations)
            {
                var filename = Path.GetFileNameWithoutExtension(project.Key) + ".xml";

            }
        }
    }
}