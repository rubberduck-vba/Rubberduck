﻿using System.Collections.Generic;
using System.IO.Abstractions;
using System.Linq;
using Rubberduck.Interaction;
using Rubberduck.InternalApi.Extensions;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class ReplaceProjectContentsFromFilesCommand : ImportCommand
    {
        public ReplaceProjectContentsFromFilesCommand(
            IVBE vbe, 
            IFileSystemBrowserFactory dialogFactory, 
            IVbeEvents vbeEvents, 
            IParseManager parseManager,
            IDeclarationFinderProvider declarationFinderProvider,
            IProjectsProvider projectsProvider,
            IModuleNameFromFileExtractor moduleNameFromFileExtractor,
            IEnumerable<IRequiredBinaryFilesFromFileNameExtractor> binaryFileExtractors,
            IFileSystem fileSystem,
            IMessageBox messageBox,
            IConfigurationService<ProjectSettings> projectSettingsProvider) 
            :base(vbe, dialogFactory, vbeEvents, parseManager, declarationFinderProvider, projectsProvider, moduleNameFromFileExtractor, binaryFileExtractors, fileSystem, messageBox, projectSettingsProvider)
        {}

        protected override string DialogsTitle => RubberduckUI.ReplaceProjectContentsFromFilesCommand_DialogCaption;

        protected override ICollection<QualifiedModuleName> ModulesToRemoveBeforeImport(IDictionary<string, QualifiedModuleName> existingModules)
        {
            return DeclarationFinderProvider.DeclarationFinder
                .UserDeclarations(DeclarationType.Module)
                .Select(decl => decl.QualifiedModuleName)
                .ToHashSet();
        }

        protected override bool UserDeclinesExecution(IVBProject targetProject)
        {
            return !UserConfirmsToReplaceProjectContents(targetProject);
        }

        private bool UserConfirmsToReplaceProjectContents(IVBProject project)
        {
            var projectName = project.Name;
            var message = string.Format(RubberduckUI.ReplaceProjectContentsFromFilesCommand_DialogCaption, projectName);
            return MessageBox.ConfirmYesNo(message, DialogsTitle, false);
        }
    }
}
