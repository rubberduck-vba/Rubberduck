using System.Linq;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.CodeAnalysis.Settings;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.SettingsProvider;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Adds an identifier or Hungarian Notation prefix to a list of white-listed identifiers and prefixes in Rubberduck's inspection settings.
    /// </summary>
    /// <inspections>
    /// <inspection name="HungarianNotationInspection" />
    /// <inspection name="UseMeaningfulNameInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="false" module="false" project="false" all="false"/>
    internal sealed class AddIdentifierToWhiteListQuickFix : QuickFixBase
    {
        private readonly IConfigurationService<CodeInspectionSettings> _settings;

        public AddIdentifierToWhiteListQuickFix(IConfigurationService<CodeInspectionSettings> settings)
            : base(typeof(HungarianNotationInspection), typeof(UseMeaningfulNameInspection))
        {
            _settings = settings;
        }

        //The rewriteSession is optional since it is not used in this particular quickfix.
        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession = null)
        {
            var inspectionSettings = _settings.Read();
            var whitelist = inspectionSettings.WhitelistedIdentifiers;
            inspectionSettings.WhitelistedIdentifiers =
                whitelist.Concat(new[] { new WhitelistedIdentifierSetting(result.Target.IdentifierName) }).ToArray();
            _settings.Save(inspectionSettings);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.WhiteListIdentifierQuickFix;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => false;
        public override bool CanFixInProject => false;
        public override bool CanFixAll => false;
    }
}
