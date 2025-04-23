using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.InternalApi.Common
{
    /// <summary>
    /// Extension to StringBuilder to allow adding text line by line.
    /// </summary>
    public class StringLineBuilder
    {
        private readonly StringBuilder _document = new StringBuilder();

        public override string ToString() => _document.ToString();

        public void AppendLine(string value = "")
            => _document.Append(value + "\r\n");

        public void AppendLineNoNullChars(string value)
            => AppendLine(value.Replace("\0", string.Empty));
    }
}
