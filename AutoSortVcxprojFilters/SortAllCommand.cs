//------------------------------------------------------------------------------
// <copyright file="SortAllCommand.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Shell;

namespace AutoSortVcxprojFilters
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class SortAllCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("d01a3a77-427a-4d7f-a5f5-85f0ad7e88e8");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        private EnvDTE.DTE m_dte;

        /// <summary>
        /// Initializes a new instance of the <see cref="SortAllCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private SortAllCommand(AsyncPackage package, OleMenuCommandService commandService, EnvDTE.DTE dte)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
            m_dte = dte;
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static SortAllCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in SortEdmx's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            EnvDTE.DTE dte = await package.GetServiceAsync(typeof(EnvDTE.DTE)) as EnvDTE.DTE;
            Instance = new SortAllCommand(package, commandService, dte);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            var projects = m_dte.Solution.Projects
               .Cast<EnvDTE.Project>()
               .Where(x => { return x?.Object != null; })
               .ToArray();

            if (projects != null)
            {
                foreach (var proj in projects)
                {
                    VCXFilterSorter.Sort(proj.FullName);
                    VCXFilterSorter.Sort(proj.FullName + @".filters");
                }
            }
        }
    }
}
