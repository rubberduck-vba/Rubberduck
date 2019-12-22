using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.Command.MenuItems.CommandBars
{
    public class SerializeProjectsCommandMenuItem : CommandMenuItemBase
    {
        public SerializeProjectsCommandMenuItem(SerializeProjectsCommand command) : base(command)
        {
        }

        public override Func<string> Caption { get { return () => "Serialize"; } }
        public override string Key => "SerializeProjects";
    }

    public class SerializeProjectsCommand : CommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IComProjectSerializationProvider _serializationProvider;
        private readonly IComLibraryProvider _comLibraryProvider;

        public SerializeProjectsCommand(RubberduckParserState state, IComProjectSerializationProvider serializationProvider, IComLibraryProvider comLibraryProvider) 
        {
            _state = state;
            _serializationProvider = serializationProvider;
            _comLibraryProvider = comLibraryProvider;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        private bool SpecialEvaluateCanExecute(object parameter)
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
                using (var references = project.References)
                {
                    foreach (var reference in references)
                    {
                        var info = new ReferenceInfo(reference);
                        reference.Dispose();
                        var library = _comLibraryProvider.LoadTypeLibrary(info.FullPath);
                        if (library == null)
                        {
                            Logger.Warn($"Could not load library {info.FullPath} for serialization.");
                            continue;
                        }

                        var type = new ComProject(library, info.FullPath);
                        if (!toSerialize.ContainsKey(type.Guid))
                        {
                            toSerialize.Add(type.Guid, type);
                        }
                    }
                }
            }

            foreach (var library in toSerialize.Values)
            {
                Logger.Warn($"Serializing {library.Path}.");
                _serializationProvider.SerializeProject(library);
            }

            SerializeComSafe();
        }

        [Conditional("TRACE_COM_SAFE")]
        private void SerializeComSafe()
        {
            //This block must be inside a conditional compilation block because the Serialize method 
            //called is conditionally compiled and available only if the compilation constant TRACE_COM_SAFE is set.
            var path = !string.IsNullOrWhiteSpace(_serializationProvider.Target)
                ? Path.GetDirectoryName(_serializationProvider.Target)
                : Path.GetTempPath();
            var traceDirectory = Path.Combine(path, "COM Trace");
            if (!Directory.Exists(traceDirectory))
            {
                Directory.CreateDirectory(traceDirectory);
            }

            Rubberduck.VBEditor.ComManagement.ComSafeManager.GetCurrentComSafe().Serialize(traceDirectory);
        }
    }
}