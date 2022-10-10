using Rubberduck.UI.Controls;
using Rubberduck.UI.FindSymbol;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Media;

namespace Rubberduck.UI.Converters
{
    /// <summary>
    /// A converter that highlights the search terms in the  a <see cref="SearchResultItem"/>.
    /// </summary>
    /// <remarks>
    /// Based on https://stackoverflow.com/a/22026985/1188513
    /// </remarks>
    class SearchResultToXamlConverter : IValueConverter
    {

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            const char nonBreakingSpace = '\u00A0';
            if (value is SearchResultItem item)
            {
                var textBlock = new TextBlock();
                textBlock.TextWrapping = TextWrapping.Wrap;

                var input = item.ResultText.Replace(' ', nonBreakingSpace);
                if (item.HighlightIndex.HasValue
                    && item.HighlightIndex.Value.EndColumn < input.Length   // if we do not check this, any inconsistent input will crash the host.
                    && item.HighlightIndex.Value.StartColumn < input.Length)
                {
                    var highlight = item.HighlightIndex.Value;
                    if (highlight.StartColumn > 0) 
                    {
                        var preRun = new Run(input.Substring(0, highlight.StartColumn))
                        {
                            Foreground = Brushes.DimGray,
                            FontFamily = new FontFamily("Consolas")
                        };
                        textBlock.Inlines.Add(preRun);
                    }

                    var highlightRun = new Run(input.Substring(highlight.StartColumn, 
                        highlight.EndLine == highlight.StartLine 
                                ? highlight.EndColumn - highlight.StartColumn
                                : highlight.StartColumn + highlight.EndColumn - 1))
                    {
                        Background = Brushes.Yellow,
                        Foreground = Brushes.DimGray,
                        FontWeight = FontWeights.Bold,
                        FontFamily = new FontFamily("Consolas")
                    };
                    textBlock.Inlines.Add(highlightRun);

                    if (highlight.EndColumn < item.ResultText.Length - 1)
                    {
                        var postRun = new Run(input.Substring(highlight.EndColumn))
                        {
                            Foreground = Brushes.DimGray,
                            FontFamily = new FontFamily("Consolas")
                        };
                        textBlock.Inlines.Add(postRun);
                    }
                }
                else
                {
                    textBlock.Inlines.Add(new Run(input));
                }

                return textBlock;
            }

            return null;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException("This converter cannot be used in two-way binding.");
        }

    }
}
