using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.UI;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class HungarianNotationInspection : InspectionBase
    {
        #region statics
        private static readonly List<string> HungarianPrefixes = new List<string>
        {
            "chk",
            "cbo",
            "cmd",
            "btn",
            "fra",
            "img",
            "lbl",
            "lst",
            "mnu",
            "opt",
            "pic",
            "shp",
            "txt",
            "tmr",
            "chk",
            "dlg",
            "drv",
            "frm",
            "grd",
            "obj",
            "rpt",
            "fld",
            "idx",
            "tbl",
            "tbd",
            "bas",
            "cls",
            "g",
            "m",
            "bln",
            "byt",
            "col",
            "dtm",
            "dbl",
            "cur",
            "int",
            "lng",
            "sng",
            "str",
            "udt",
            "vnt",
            "var",
            "pgr",
            "dao",
            "b",
            "by",
            "c",
            "chr",
            "i",
            "l",
            "s",
            "o",
            "n",
            "dt",
            "dat",
            "a",
            "arr"
        };

        private static readonly Regex HungarianIdentifierRegex = new Regex($"^({string.Join("|", HungarianPrefixes)})[A-Z0-9].*$");

        private static readonly List<DeclarationType> TargetDeclarationTypes = new List<DeclarationType>
        {
            DeclarationType.Parameter,
            DeclarationType.Constant,
            DeclarationType.Control,
            DeclarationType.ClassModule,
            DeclarationType.Member,
            DeclarationType.Module,
            DeclarationType.ProceduralModule,
            DeclarationType.UserForm,
            DeclarationType.UserDefinedType,
            DeclarationType.UserDefinedTypeMember,
            DeclarationType.Variable
        };

        #endregion
        
        private readonly IPersistanceService<CodeInspectionSettings> _settings;

        public HungarianNotationInspection(RubberduckParserState state, IPersistanceService<CodeInspectionSettings> settings)
            : base(state, CodeInspectionSeverity.Suggestion)
        {
            _settings = settings;
        }

        public override CodeInspectionType InspectionType => CodeInspectionType.MaintainabilityAndReadabilityIssues;

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            var settings = _settings.Load(new CodeInspectionSettings()) ?? new CodeInspectionSettings();
            var whitelistedNames = settings.WhitelistedIdentifiers.Select(s => s.Identifier).ToList();

            var hungarians = UserDeclarations
                .Where(declaration => !whitelistedNames.Contains(declaration.IdentifierName) &&
                                      TargetDeclarationTypes.Contains(declaration.DeclarationType) &&
                                      HungarianIdentifierRegex.IsMatch(declaration.IdentifierName))
                .Select(issue => new DeclarationInspectionResult(this,
                                                      string.Format(InspectionsUI.IdentifierNameInspectionResultFormat,
                                                                    RubberduckUI.ResourceManager.GetString("DeclarationType_" + issue.DeclarationType, CultureInfo.CurrentUICulture),
                                                                    issue.IdentifierName),
                                                      issue));

            return hungarians;
        }
    }
}
