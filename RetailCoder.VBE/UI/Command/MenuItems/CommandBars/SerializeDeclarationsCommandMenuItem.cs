using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Printing;
using System.Runtime.Serialization;
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
        private readonly IPersistable<SerializableDeclarationTree> _service;

        public SerializeDeclarationsCommand(RubberduckParserState state, IPersistable<SerializableDeclarationTree> service) 
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

            foreach (var tree in _state.BuiltInDeclarationTrees)
            {
                System.Diagnostics.Debug.Assert(path != null, "project path isn't supposed to be null");

                var filename = Path.GetFileNameWithoutExtension(tree.Node.QualifiedMemberName.QualifiedModuleName.ProjectName) + ".xml";
                _service.Persist(Path.Combine(path, filename), tree);
            }
        }
    }

    public class XmlPersistableDeclarations : IPersistable<SerializableDeclarationTree>
    {
        public void Persist(string path, SerializableDeclarationTree tree)
        {
            if (string.IsNullOrEmpty(path)) { throw new InvalidOperationException(); }

            var xmlSettings = new XmlWriterSettings
            {
                NamespaceHandling = NamespaceHandling.OmitDuplicates,
                Encoding = Encoding.UTF8,
                //Indent = true
            };

            using (var stream = new FileStream(path, FileMode.Create, FileAccess.Write))
            using (var xmlWriter = XmlWriter.Create(stream, xmlSettings))
            using (var writer = XmlDictionaryWriter.CreateDictionaryWriter(xmlWriter))
            {
                writer.WriteStartDocument();
                var settings = new DataContractSerializerSettings {RootNamespace = XmlDictionaryString.Empty};
                var serializer = new DataContractSerializer(typeof (SerializableDeclarationTree), settings);
                serializer.WriteObject(writer, tree);
            }
        }

        public SerializableDeclarationTree Load(string path)
        {
            if (string.IsNullOrEmpty(path)) { throw new InvalidOperationException(); }
            throw new NotImplementedException();
        }
    }
}