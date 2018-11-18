using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.ComManagement;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class PrintCommand : CommandBase
    {
        private readonly IProjectsProvider _projectsProvider;

        public PrintCommand(IProjectsProvider projectsProvider)
            : base(LogManager.GetCurrentClassLogger())
        {
            _projectsProvider = projectsProvider;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            var node = parameter as CodeExplorerComponentViewModel;
            if (node == null)
            {
                return false;
            }

            try
            {
                var component = _projectsProvider.Component(node.Declaration.QualifiedName.QualifiedModuleName);
                using (var codeModule = component.CodeModule)
                {
                    return codeModule.CountOfLines != 0;
                }
            }
            catch (COMException)
            {
                // thrown when the component reference is stale
                return false;
            }
        }

        protected override void OnExecute(object parameter)
        {
            var node = (CodeExplorerComponentViewModel)parameter;
            var qualifiedComponentName = node.Declaration.QualifiedName.QualifiedModuleName;

            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Rubberduck",
                qualifiedComponentName.ComponentName + ".txt");

            List<string> text;
            var component = _projectsProvider.Component(qualifiedComponentName);
            using (var codeModule = component.CodeModule)
            {
                text = codeModule.GetLines(1, codeModule.CountOfLines)
                    .Split(new[] {Environment.NewLine}, StringSplitOptions.None).ToList();
            }

            var printDoc = new PrintDocument { DocumentName = path };
            using (var pd = new PrintDialog
            {
                Document = printDoc,
                AllowCurrentPage = true,
                AllowSelection = true,
                AllowPrintToFile = true,
                AllowSomePages = true,
                UseEXDialog = true
            })
            {
                if (pd.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                var offsetY = 0;
                var pageHeight = pd.PrinterSettings.PaperSizes[0].Height;

                var index = 0;

                printDoc.PrintPage += (sender, printPageArgs) =>
                {
                    while (index < text.Count)
                    {
                        using (var font = new Font(new FontFamily("Consolas"), 10, FontStyle.Regular))
                        using (var stringFormat = new StringFormat())
                        {
                            printPageArgs.Graphics.DrawString(text[index++], font, Brushes.Black, 0, offsetY,
                                stringFormat);

                            offsetY += font.Height;

                            if (offsetY >= pageHeight)
                            {
                                printPageArgs.HasMorePages = true;
                                offsetY = 0;
                                return;
                            }

                            printPageArgs.HasMorePages = false;
                        }
                    }
                };

                printDoc.Print();
            }
        }
    }
}
