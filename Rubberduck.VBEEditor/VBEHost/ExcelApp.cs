﻿using Microsoft.Office.Interop.Excel;

namespace Rubberduck.VBEditor.VBEHost
{
    public class ExcelApp : HostApplicationBase<Application>
    {
        public ExcelApp() : base("Excel") { }

        public override void Run(QualifiedMemberName qualifiedMemberName)
        {
            Application.Run(GenerateMethodCall(qualifiedMemberName));
        }

        protected virtual string GenerateMethodCall(QualifiedMemberName qualifiedMemberName)
        {
            return qualifiedMemberName.ToString();
        }
    }
}