using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NLog;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.Command.MenuItems.CommandBars
{
    public class SerializeProjectsCommandMenuItem : CommandMenuItemBase
    {
        public SerializeProjectsCommandMenuItem(SerializeDeclarationsCommand command) : base(command)
        {
        }

        public override Func<string> Caption { get { return () => "Serialize"; } }
        public override string Key => "SerializeProjects";
    }

    public class SerializeDeclarationsCommand : CommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IComProjectSerializationProvider _serializationProvider;
        private readonly IComLibraryProvider _comLibraryProvider;

        public SerializeDeclarationsCommand(RubberduckParserState state, IComProjectSerializationProvider serializationProvider, IComLibraryProvider comLibraryProvider) 
            : base(LogManager.GetCurrentClassLogger())
        {
            _state = state;
            _serializationProvider = serializationProvider;
            _comLibraryProvider = comLibraryProvider;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            return _state.Status == ParserState.Ready;
        }

        protected override void OnExecute(object parameter)
        {
            if (!Directory.Exists(_serializationProvider.Target))
            {
                Directory.CreateDirectory(_serializationProvider.Target);
            }

            var toSerialize = new Dictionary<Guid, ComProject>();

            foreach (var project in _state.ProjectsProvider.Projects().Select(proj => proj.Project))
            {
                foreach (var reference in project.References.Select(lib => new ReferenceInfo(lib)))
                {
                    var library = _comLibraryProvider.LoadTypeLibrary(reference.FullPath);
                    if (library == null)
                    {
                        Logger.Warn($"Could not load library {reference.FullPath} for serialization.");
                        continue;
                    }

                    var type = new ComProject(library, reference.FullPath);
                    if (!toSerialize.ContainsKey(type.Guid))
                    {
                        toSerialize.Add(type.Guid, type);
                    }
                }
            }

            foreach (var library in toSerialize.Values)
            {
                Logger.Warn($"Serializing {library.Path}.");
                _serializationProvider.SerializeProject(library);
            }
        }
    }
}