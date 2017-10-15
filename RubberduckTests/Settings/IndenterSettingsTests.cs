using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Settings;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.Settings;

namespace RubberduckTests.Settings
{
    [TestClass]
    public class IndenterSettingsTests
    {
        #region Defaults
        private const bool DefaultIndentEntireProcedureBody = true;
        private const bool DefaultIndentFirstCommentBlock = true;
        private const bool DefaultIndentFirstDeclarationBlock = true;
        private const bool DefaultAlignCommentsWithCode = true;
        private const bool DefaultAlignContinuations = true;
        private const bool DefaultIgnoreOperatorsInContinuations = true;
        private const bool DefaultIndentCase = false;
        private const bool DefaultIndentEnumTypeAsProcedure = false;
        private const bool DefaultForceDebugStatementsInColumn1 = false;
        private const bool DefaultForceCompilerDirectivesInColumn1 = false;
        private const bool DefaultIndentCompilerDirectives = true;
        private const bool DefaultAlignDims = false;
        private const int DefaultAlignDimColumn = 15;
        private const EndOfLineCommentStyle DefaultEndOfLineCommentStyle = EndOfLineCommentStyle.AlignInColumn;
        private const int DefaultEndOfLineCommentColumnSpaceAlignment = 50;
        private const int DefaultIndentSpaces = 4;
        private const bool DefaultVerticallySpaceProcedures = true;
        private const int DefaultLinesBetweenProcedures = 1;
        #endregion

        #region Nondefaults
        private const bool NondefaultIndentEntireProcedureBody = false;
        private const bool NondefaultIndentFirstCommentBlock = false;
        private const bool NondefaultIndentFirstDeclarationBlock = false;
        private const bool NondefaultAlignCommentsWithCode = false;
        private const bool NondefaultAlignContinuations = false;
        private const bool NondefaultIgnoreOperatorsInContinuations = false;
        private const bool NondefaultIndentCase = true;
        private const bool NondefaultIndentEnumTypeAsProcedure = true;
        private const bool NondefaultForceDebugStatementsInColumn1 = true;
        private const bool NondefaultForceCompilerDirectivesInColumn1 = true;
        private const bool NondefaultIndentCompilerDirectives = false;
        private const bool NondefaultAlignDims = true;
        private const int NondefaultAlignDimColumn = 16;
        private const EndOfLineCommentStyle NondefaultEndOfLineCommentStyle = EndOfLineCommentStyle.Absolute;
        private const int NondefaultEndOfLineCommentColumnSpaceAlignment = 60;
        private const int NondefaultIndentSpaces = 2;
        private const bool NondefaultVerticallySpaceProcedures = false;
        private const int NondefaultLinesBetweenProcedures = 2;
        #endregion

        public static Rubberduck.SmartIndenter.IndenterSettings GetMockIndenterSettings(bool nondefault = false)
        {
            var output = new Mock<Rubberduck.SmartIndenter.IndenterSettings>();

            output.SetupProperty(s => s.IndentEntireProcedureBody);
            output.SetupProperty(s => s.IndentFirstCommentBlock);
            output.SetupProperty(s => s.IndentFirstDeclarationBlock);
            output.SetupProperty(s => s.AlignCommentsWithCode);
            output.SetupProperty(s => s.AlignContinuations);
            output.SetupProperty(s => s.IgnoreOperatorsInContinuations);
            output.SetupProperty(s => s.IndentCase);
            output.SetupProperty(s => s.IndentEnumTypeAsProcedure);
            output.SetupProperty(s => s.ForceDebugStatementsInColumn1);
            output.SetupProperty(s => s.ForceCompilerDirectivesInColumn1);
            output.SetupProperty(s => s.IndentCompilerDirectives);
            output.SetupProperty(s => s.AlignDims);
            output.SetupProperty(s => s.AlignDimColumn);
            output.SetupProperty(s => s.EndOfLineCommentStyle);
            output.SetupProperty(s => s.EndOfLineCommentColumnSpaceAlignment);
            output.SetupProperty(s => s.IndentSpaces);
            output.SetupProperty(s => s.VerticallySpaceProcedures);
            output.SetupProperty(s => s.LinesBetweenProcedures);

            output.Object.IndentEntireProcedureBody = nondefault ? NondefaultIndentEntireProcedureBody : DefaultIndentEntireProcedureBody;
            output.Object.IndentFirstCommentBlock = nondefault ? NondefaultIndentFirstCommentBlock : DefaultIndentFirstCommentBlock;
            output.Object.IndentFirstDeclarationBlock = nondefault ? NondefaultIndentFirstDeclarationBlock : DefaultIndentFirstDeclarationBlock;
            output.Object.AlignCommentsWithCode = nondefault ? NondefaultAlignCommentsWithCode : DefaultAlignCommentsWithCode;
            output.Object.AlignContinuations = nondefault ? NondefaultAlignContinuations : DefaultAlignContinuations;
            output.Object.IgnoreOperatorsInContinuations = nondefault ? NondefaultIgnoreOperatorsInContinuations : DefaultIgnoreOperatorsInContinuations;
            output.Object.IndentCase = nondefault ? NondefaultIndentCase : DefaultIndentCase;
            output.Object.IndentEnumTypeAsProcedure = nondefault ? NondefaultIndentEnumTypeAsProcedure : DefaultIndentEnumTypeAsProcedure;
            output.Object.ForceDebugStatementsInColumn1 = nondefault ? NondefaultForceDebugStatementsInColumn1 : DefaultForceDebugStatementsInColumn1;
            output.Object.ForceCompilerDirectivesInColumn1 = nondefault ? NondefaultForceCompilerDirectivesInColumn1 : DefaultForceCompilerDirectivesInColumn1;
            output.Object.IndentCompilerDirectives = nondefault ? NondefaultIndentCompilerDirectives : DefaultIndentCompilerDirectives;
            output.Object.AlignDims = nondefault ? NondefaultAlignDims : DefaultAlignDims;
            output.Object.AlignDimColumn = nondefault ? NondefaultAlignDimColumn : DefaultAlignDimColumn;
            output.Object.EndOfLineCommentStyle = nondefault ? NondefaultEndOfLineCommentStyle : DefaultEndOfLineCommentStyle;
            output.Object.EndOfLineCommentColumnSpaceAlignment = nondefault ? NondefaultEndOfLineCommentColumnSpaceAlignment : DefaultEndOfLineCommentColumnSpaceAlignment;
            output.Object.IndentSpaces = nondefault ? NondefaultIndentSpaces : DefaultIndentSpaces;
            output.Object.VerticallySpaceProcedures = nondefault ? NondefaultVerticallySpaceProcedures : DefaultVerticallySpaceProcedures;
            output.Object.LinesBetweenProcedures = nondefault ? NondefaultLinesBetweenProcedures : DefaultLinesBetweenProcedures;

            return output.Object;
        }

        private Configuration GetDefaultConfig()
        {
            var userSettings = new UserSettings(null, null, null, null, null, GetMockIndenterSettings(), null);
            return new Configuration(userSettings);
        }

        private Configuration GetNondefaultConfig()
        {
            var userSettings = new UserSettings(null, null, null, null, null, GetMockIndenterSettings(true), null);
            return new Configuration(userSettings);
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void SaveConfigWorks()
        {
            var customConfig = GetNondefaultConfig();
            var viewModel = new IndenterSettingsViewModel(customConfig);

            var config = GetDefaultConfig();
            viewModel.UpdateConfig(config);

            MultiAssert.Aggregate(
                () => Assert.AreEqual(config.UserSettings.IndenterSettings.AlignCommentsWithCode, viewModel.AlignCommentsWithCode),
                () => Assert.AreEqual(config.UserSettings.IndenterSettings.AlignContinuations, viewModel.AlignContinuations),
                () => Assert.AreEqual(config.UserSettings.IndenterSettings.AlignDimColumn, viewModel.AlignDimColumn),
                () => Assert.AreEqual(config.UserSettings.IndenterSettings.AlignDims, viewModel.AlignDims),
                () => Assert.AreEqual(config.UserSettings.IndenterSettings.EndOfLineCommentColumnSpaceAlignment, viewModel.EndOfLineCommentColumnSpaceAlignment),
                () => Assert.AreEqual(config.UserSettings.IndenterSettings.EndOfLineCommentStyle, viewModel.EndOfLineCommentStyle),
                () => Assert.AreEqual(config.UserSettings.IndenterSettings.ForceCompilerDirectivesInColumn1, viewModel.ForceCompilerDirectivesInColumn1),
                () => Assert.AreEqual(config.UserSettings.IndenterSettings.ForceDebugStatementsInColumn1, viewModel.ForceDebugStatementsInColumn1),
                () => Assert.AreEqual(config.UserSettings.IndenterSettings.IgnoreOperatorsInContinuations, viewModel.IgnoreOperatorsInContinuations),
                () => Assert.AreEqual(config.UserSettings.IndenterSettings.IndentCase, viewModel.IndentCase),
                () => Assert.AreEqual(config.UserSettings.IndenterSettings.IndentEnumTypeAsProcedure, viewModel.IndentEnumTypeAsProcedure),
                () => Assert.AreEqual(config.UserSettings.IndenterSettings.IndentCompilerDirectives, viewModel.IndentCompilerDirectives),
                () => Assert.AreEqual(config.UserSettings.IndenterSettings.IndentEntireProcedureBody, viewModel.IndentEntireProcedureBody),
                () => Assert.AreEqual(config.UserSettings.IndenterSettings.IndentFirstCommentBlock, viewModel.IndentFirstCommentBlock),
                () => Assert.AreEqual(config.UserSettings.IndenterSettings.IndentFirstDeclarationBlock, viewModel.IndentFirstDeclarationBlock),
                () => Assert.AreEqual(config.UserSettings.IndenterSettings.IndentSpaces, viewModel.IndentSpaces),
                () => Assert.AreEqual(config.UserSettings.IndenterSettings.VerticallySpaceProcedures, viewModel.VerticallySpaceProcedures),
                () => Assert.AreEqual(config.UserSettings.IndenterSettings.LinesBetweenProcedures, viewModel.LinesBetweenProcedures));
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void SetDefaultsWorks()
        {
            var viewModel = new IndenterSettingsViewModel(GetNondefaultConfig());

            var defaultConfig = GetDefaultConfig();
            viewModel.SetToDefaults(defaultConfig);

            MultiAssert.Aggregate(
                () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.AlignCommentsWithCode, viewModel.AlignCommentsWithCode),
                () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.AlignContinuations, viewModel.AlignContinuations),
                () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.AlignDimColumn, viewModel.AlignDimColumn),
                () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.AlignDims, viewModel.AlignDims),
                () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.EndOfLineCommentColumnSpaceAlignment, viewModel.EndOfLineCommentColumnSpaceAlignment),
                () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.EndOfLineCommentStyle, viewModel.EndOfLineCommentStyle),
                () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.ForceCompilerDirectivesInColumn1, viewModel.ForceCompilerDirectivesInColumn1),
                () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.ForceDebugStatementsInColumn1, viewModel.ForceDebugStatementsInColumn1),
                () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IgnoreOperatorsInContinuations, viewModel.IgnoreOperatorsInContinuations),
                () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentCase, viewModel.IndentCase),
                () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentEnumTypeAsProcedure, viewModel.IndentEnumTypeAsProcedure),
                () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentCompilerDirectives, viewModel.IndentCompilerDirectives),
                () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentEntireProcedureBody, viewModel.IndentEntireProcedureBody),
                () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentFirstCommentBlock, viewModel.IndentFirstCommentBlock),
                () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentFirstDeclarationBlock, viewModel.IndentFirstDeclarationBlock),
                () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentSpaces, viewModel.IndentSpaces),
                () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.VerticallySpaceProcedures, viewModel.VerticallySpaceProcedures),
                () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.LinesBetweenProcedures, viewModel.LinesBetweenProcedures));
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void AlignCommentsWithCodeIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.AlignCommentsWithCode, viewModel.AlignCommentsWithCode);
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void AlignContinuationsIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.AlignContinuations, viewModel.AlignContinuations);
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void AlignDimColumnIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.AlignDimColumn, viewModel.AlignDimColumn);
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void AlignDimsIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.AlignDims, viewModel.AlignDims);
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void EndOfLineCommentColumnSpaceAlignmentIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.EndOfLineCommentColumnSpaceAlignment, viewModel.EndOfLineCommentColumnSpaceAlignment);
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void EndOfLineCommentStyleIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.EndOfLineCommentStyle, viewModel.EndOfLineCommentStyle);
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void ForceCompilerDirectivesInColumn1IsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.ForceCompilerDirectivesInColumn1, viewModel.ForceCompilerDirectivesInColumn1);
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void ForceDebugStatementsInColumn1IsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.ForceDebugStatementsInColumn1, viewModel.ForceDebugStatementsInColumn1);
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void IgnoreOperatorsInContinuationsIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IgnoreOperatorsInContinuations, viewModel.IgnoreOperatorsInContinuations);
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void IndentCaseIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentCase, viewModel.IndentCase);
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void IndentEnumTypeAsProcedureIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentEnumTypeAsProcedure, viewModel.IndentEnumTypeAsProcedure);
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void IndentCompilerDirectivesIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentCompilerDirectives, viewModel.IndentCompilerDirectives);
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void IndentEntireProcedureBodyIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentEntireProcedureBody, viewModel.IndentEntireProcedureBody);
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void IndentFirstCommentBlockIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentFirstCommentBlock, viewModel.IndentFirstCommentBlock);
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void IndentFirstDeclarationBlockIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentFirstDeclarationBlock, viewModel.IndentFirstDeclarationBlock);
        }

        [TestCategory("Settings")]
        [TestMethod]
        public void VerticallySpaceProceduresIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.VerticallySpaceProcedures, viewModel.VerticallySpaceProcedures);
        }

        [TestMethod]
        public void LinesBetweenProceduresIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.LinesBetweenProcedures, viewModel.LinesBetweenProcedures);
        }
    }
}
