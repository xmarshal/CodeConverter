using System;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ICSharpCode.CodeConverter.CSharp;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Threading;
using IAsyncServiceProvider = Microsoft.VisualStudio.Shell.IAsyncServiceProvider;
using OleMenuCommand = Microsoft.VisualStudio.Shell.OleMenuCommand;
using OleMenuCommandService = Microsoft.VisualStudio.Shell.OleMenuCommandService;
using Task = System.Threading.Tasks.Task;

namespace CodeConverter.VsExtension
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class ConvertVBToCSCommand
    {
        public const int MainMenuCommandId = 0x0200;
        public const int CtxMenuCommandId = 0x0201;
        public const int ProjectItemCtxMenuCommandId = 0x0202;
        public const int SolutionOrProjectCtxMenuCommandId = 0x0203;
        private const string ProjectExtension = ".vbproj";

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("a3378a21-e939-40c9-9e4b-eb0cec7b7854");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        readonly REConverterPackage _package;

        private CodeConversion _codeConversion;

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static ConvertVBToCSCommand Instance {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        IAsyncServiceProvider ServiceProvider => _package.AsyncServiceProvider;

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(REConverterPackage package)
        {
            var oleMenuCommandService = package.AsyncServiceProvider.GetServiceAsync(typeof(IMenuCommandService));
            Instance = new ConvertVBToCSCommand(package);
            await Instance.InitializeAsync(async () => (OleMenuCommandService)await oleMenuCommandService);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ConvertVBToCSCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        ConvertVBToCSCommand(REConverterPackage package)
        {
            this._package = package ?? throw new ArgumentNullException(nameof(package));
            _codeConversion = new CodeConversion(package.AsyncServiceProvider, package.GetWorkspaceAsync, () => package.Options);

        }

        private async Task InitializeAsync(Func<Task<OleMenuCommandService>> commandServiceTask)
        {
            // Command in main menu
            var menuCommandId = new CommandID(CommandSet, MainMenuCommandId);
            var menuItem = new OleMenuCommand(CodeEditorMenuItemCallbackAsync, menuCommandId);
            menuItem.BeforeQueryStatus += CodeEditorMenuItem_BeforeQueryStatus;

            // Command in code editor's context menu
            var ctxMenuCommandId = new CommandID(CommandSet, CtxMenuCommandId);
            var ctxMenuItem = new OleMenuCommand(CodeEditorMenuItemCallbackAsync, ctxMenuCommandId);
            ctxMenuItem.BeforeQueryStatus += CodeEditorMenuItem_BeforeQueryStatus;

            // Command in project item context menu
            var projectItemCtxMenuCommandId = new CommandID(CommandSet, ProjectItemCtxMenuCommandId);
            var projectItemCtxMenuItem = new OleMenuCommand(ProjectItemMenuItemCallbackAsync, projectItemCtxMenuCommandId);
            projectItemCtxMenuItem.BeforeQueryStatus += ProjectItemMenuItem_BeforeQueryStatus;

            // Command in project context menu
            var solutionOrProjectCtxMenuCommandId = new CommandID(CommandSet, SolutionOrProjectCtxMenuCommandId);
            var solutionOrProjectCtxMenuItem =
                new OleMenuCommand(SolutionOrProjectMenuItemCallbackAsync, solutionOrProjectCtxMenuCommandId);
            solutionOrProjectCtxMenuItem.BeforeQueryStatus += SolutionOrProjectMenuItem_BeforeQueryStatus;

            var commandService = await commandServiceTask();
            if (commandService != null) {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                commandService.AddCommand(menuItem);
                commandService.AddCommand(ctxMenuItem);
                commandService.AddCommand(projectItemCtxMenuItem);
                commandService.AddCommand(solutionOrProjectCtxMenuItem);
                await TaskScheduler.Default;
            }
        }

        void CodeEditorMenuItem_BeforeQueryStatus(object sender, EventArgs e)
        {
            var menuItem = sender as OleMenuCommand;
            if (menuItem != null) {

                var selection = ThreadHelper.JoinableTaskFactory.Run(async () =>
                    await _codeConversion.GetSelectionInCurrentViewAsync(CodeConversion.IsVBFileName));
                menuItem.Visible = !selection?.StreamSelectionSpan.IsEmpty ?? false;
            }
        }

        void ProjectItemMenuItem_BeforeQueryStatus(object sender, EventArgs e)
        {
            var menuItem = sender as OleMenuCommand;
            if (menuItem != null) {
                menuItem.Visible = false;
                menuItem.Enabled = false;

                string itemPath = VisualStudioInteraction.GetSingleSelectedItemOrDefault()?.ItemPath;
                if (itemPath == null || !CodeConversion.IsVBFileName(itemPath))
                    return;

                menuItem.Visible = true;
                menuItem.Enabled = true;
            }
        }

        private void SolutionOrProjectMenuItem_BeforeQueryStatus(object sender, EventArgs e)
        {
            var menuItem = sender as OleMenuCommand;
            if (menuItem != null) {
                menuItem.Visible = menuItem.Enabled = VisualStudioInteraction.GetSelectedProjects(ProjectExtension).Any();
            }
        }

#pragma warning disable AvoidAsyncVoid // Avoid async void
        async void CodeEditorMenuItemCallbackAsync(object sender, EventArgs e)
#pragma warning restore AvoidAsyncVoid // Avoid async void
        {
            var span = (await _codeConversion.GetSelectionInCurrentViewAsync(CodeConversion.IsVBFileName)).SelectedSpans.First().Span;
            await ConvertDocumentAsync((await _codeConversion.GetCurrentViewHostAsync(CodeConversion.IsVBFileName)).GetTextDocument().FilePath, span);
        }

#pragma warning disable AvoidAsyncVoid // Avoid async void
        async void ProjectItemMenuItemCallbackAsync(object sender, EventArgs e)
#pragma warning restore AvoidAsyncVoid // Avoid async void
        {
            string itemPath = VisualStudioInteraction.GetSingleSelectedItemOrDefault()?.ItemPath;
            await ConvertDocumentAsync(itemPath, new Span(0, 0));
        }

#pragma warning disable AvoidAsyncVoid // Avoid async void
        private async void SolutionOrProjectMenuItemCallbackAsync(object sender, EventArgs e)
#pragma warning restore AvoidAsyncVoid // Avoid async void
        {
            try {
                var projects = VisualStudioInteraction.GetSelectedProjects(ProjectExtension);
                await _codeConversion.PerformProjectConversionAsync<VBToCSConversion>(projects);
            } catch (Exception ex) {
                VisualStudioInteraction.ShowException((IServiceProvider) await ServiceProvider.GetServiceAsync(typeof(IServiceProvider)), CodeConversion.ConverterTitle, ex);
            }
        }

        private async Task ConvertDocumentAsync(string documentPath, Span selected)
        {
            if (documentPath == null || !CodeConversion.IsVBFileName(documentPath))
                return;

            try {
                await _codeConversion.PerformDocumentConversionAsync<VBToCSConversion>(documentPath, selected);
            }
            catch (Exception ex) {
                VisualStudioInteraction.ShowException((IServiceProvider)await ServiceProvider.GetServiceAsync(typeof(IServiceProvider)), CodeConversion.ConverterTitle, ex);
            }
        }
    }
}
