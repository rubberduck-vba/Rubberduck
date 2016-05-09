using System;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CodeExplorer_PrintCommand : CommandBase
    {
        public override bool CanExecute(object parameter)
        {
            var node = parameter as CodeExplorerComponentViewModel;
            if (node == null)
            {
                return false;
            }

            return node.Declaration.QualifiedName.QualifiedModuleName.Component.CodeModule.CountOfLines != 0;
        }

        public override void Execute(object parameter)
        {
            var node = (CodeExplorerComponentViewModel)parameter;
            var component = node.Declaration.QualifiedName.QualifiedModuleName.Component;

            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Rubberduck",
                component.Name + ".txt");

            var text = component.CodeModule.Lines[1, component.CodeModule.CountOfLines].Split(new[] {Environment.NewLine}, StringSplitOptions.None).ToList();

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
                        var font = new Font(new FontFamily("Consolas"), 10, FontStyle.Regular);
                        printPageArgs.Graphics.DrawString(text[index++], font, Brushes.Black, 0, offsetY, new StringFormat());

                        offsetY += font.Height;

                        if (offsetY >= pageHeight)
                        {
                            printPageArgs.HasMorePages = true;
                            offsetY = 0;
                            return;
                        }

                        printPageArgs.HasMorePages = false;
                    }
                };

                printDoc.Print();
            }
        }
    }
}