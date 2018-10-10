using System;
using System.IO;
using System.Linq;
using NLog;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.SettingsProvider;

namespace Rubberduck.UI.Command.MenuItems.CommandBars
{
    public class SerializeDeclarationsCommandMenuItem : CommandMenuItemBase
    {
        public SerializeDeclarationsCommandMenuItem(SerializeDeclarationsCommand command) : base(command)
        {
        }

        public override Func<string> Caption { get { return () => "Serialize"; } }
        public override string Key => "SerializeDeclarations";
    }

    public class SerializeDeclarationsCommand : CommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IPersistable<SerializableProject> _service;
        private readonly ISerializableProjectBuilder _serializableProjectBuilder;

        public SerializeDeclarationsCommand(RubberduckParserState state, IPersistable<SerializableProject> service, ISerializableProjectBuilder serializableProjectBuilder) 
            : base(LogManager.GetCurrentClassLogger())
        {
            _state = state;
            _service = service;
            _serializableProjectBuilder = serializableProjectBuilder;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            return _state.Status == ParserState.Ready;
        }

        private static readonly string BasePath =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Rubberduck");

        protected override void OnExecute(object parameter)
        {
            var path = Path.Combine(BasePath, "Declarations");
            if (!Directory.Exists(path)) { Directory.CreateDirectory(path); }

            foreach (var project in _state.DeclarationFinder.BuiltInDeclarations(DeclarationType.Project).OfType<ProjectDeclaration>())
            {
                System.Diagnostics.Debug.Assert(path != null, "project path isn't supposed to be null");

                var tree = _serializableProjectBuilder.SerializableProject(project);
                var filename = $"{tree.Node.QualifiedMemberName.QualifiedModuleName.ProjectName}.{tree.MajorVersion}.{tree.MinorVersion}.xml";
                var fullFilename = Path.Combine(path, filename);
                _service.Persist(fullFilename, tree);
            }
        }
    }
}