using System;
using System.IO;
using NLog;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.SettingsProvider;

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
        private readonly IPersistable<SerializableProject> _service;

        public SerializeDeclarationsCommand(RubberduckParserState state, IPersistable<SerializableProject> service) 
            : base(LogManager.GetCurrentClassLogger())
        {
            _state = state;
            _service = service;
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            return _state.Status == ParserState.Ready;
        }

        private static readonly string BasePath =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Rubberduck");

        protected override void ExecuteImpl(object parameter)
        {
            var path = Path.Combine(BasePath, "Declarations");
            if (!Directory.Exists(path)) { Directory.CreateDirectory(path); }

            foreach (var tree in _state.BuiltInDeclarationTrees)
            {
                System.Diagnostics.Debug.Assert(path != null, "project path isn't supposed to be null");

                var filename = string.Format("{0}.{1}.{2}", tree.Node.QualifiedMemberName.QualifiedModuleName.ProjectName, tree.MajorVersion, tree.MinorVersion) + ".xml";
                _service.Persist(Path.Combine(path, filename), tree);
            }
        }
    }
}