using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Printing;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using NLog;
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
        private readonly IPersistable<SerializableDeclaration> _service;

        public SerializeDeclarationsCommand(RubberduckParserState state, IPersistable<SerializableDeclaration> service) 
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
            var path = Path.Combine(BasePath, "declarations");
            if (!Directory.Exists(path)) { Directory.CreateDirectory(path); }

            var declarations = _state.AllDeclarations
                .Where(declaration => declaration.IsBuiltIn)
                .Select(declaration => new SerializableDeclaration(declaration))
                .GroupBy(declaration => declaration.QualifiedMemberName.QualifiedModuleName.ProjectPath);
            foreach (var project in declarations)
            {
                System.Diagnostics.Debug.Assert(path != null, "project path isn't supposed to be null");

                var filename = Path.GetFileNameWithoutExtension(project.Key) + ".xml";
                _service.Persist(Path.Combine(path, filename), project);
            }
        }
    }

    public class XmlPersistableDeclarations : IPersistable<SerializableDeclaration>
    {
        public void Persist(string path, IEnumerable<SerializableDeclaration> items)
        {
            if (string.IsNullOrEmpty(path)) { throw new InvalidOperationException(); }

            var emptyNamespace = new XmlSerializerNamespaces(new[] { XmlQualifiedName.Empty });       
            using (var writer = new StreamWriter(path, false))
            {
                var serializer = new XmlSerializer(typeof(SerializableDeclaration));
                foreach (var item in items)
                {
                    serializer.Serialize(writer, item, emptyNamespace);
                }
            }
        }

        public IEnumerable<SerializableDeclaration> Load(string path)
        {
            if (string.IsNullOrEmpty(FullPath)) { throw new InvalidOperationException(); }
            throw new NotImplementedException();
        }

        public string FullPath { get; set; }
    }
}