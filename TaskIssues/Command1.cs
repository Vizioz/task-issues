﻿using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using FileNotFoundException = System.IO.FileNotFoundException;
using Process = System.Diagnostics.Process;
using EnvDTE;
using EnvDTE80;
using Microsoft;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;
using Microsoft.VisualStudio;
using System.ComponentModel;
using Microsoft.VisualStudio.Shell.TableManager;

namespace TaskIssues
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class Command1
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("26b1b304-d5bb-4c05-b2dd-c4ed9fb35619");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage _package;

        /// <summary>
        /// The current solution's GIT repository URL.
        /// </summary>
        private readonly string gitRepo;
        private readonly DTE2 _dte;
        /// <summary>
        /// Initializes a new instance of the <see cref="Command1"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private Command1(OleMenuCommandService commandService, DTE2 dte, AsyncPackage package)
        {
            _dte = dte;
            this._package = package ?? throw new ArgumentNullException(nameof(package));
            this.gitRepo = GetGitUrl();
            //commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            var commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            var dte = await package.GetServiceAsync(typeof(DTE)) as DTE2;
            Instance = new Command1(commandService, dte, package);
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
            ThreadHelper.ThrowIfNotOnUIThread();

            // Get a reference to the Development Tools Environment (DTE) - I.e. the main VS Application
            //var dte = ServiceProvider.GetService(typeof(DTE)) as DTE2;
            //Assumes.Present(dte);

            // Get the current selected task item
            var item = GetTaskItem(_dte);
            
            if (item == null)
            {
                VsShellUtilities.ShowMessageBox(
                serviceProvider: this._package,
                title: "Task Issues",
                message: "No GITHub issue was found in the task definition",
                icon: OLEMSGICON.OLEMSGICON_INFO,
                msgButton: OLEMSGBUTTON.OLEMSGBUTTON_OK,
                defaultButton: OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }
            else
            {
                // TODO: #3 - This will need to be changed to use the GITHUB API to pull in the Issue info to a custom UI within Visual Studio as GitHub does not support IE which is the embedded browser!
                // Open a browser window with the GIT issue
                //var itemOps = dte.ItemOperations;
                
                // dte.ExecuteCommand(gitRepo + item);

                OpenUri(new Uri(gitRepo + item));
            }

        }

        private bool OpenUri(Uri uri)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (uri == null)
                throw new ArgumentNullException("uri");

            if (!uri.IsAbsoluteUri)
                return false;

            /* First try to use the Web Browsing Service. This is not known to work because the
             * CreateExternalWebBrowser method always returns E_NOTIMPL. However, it is presumably
             * safer than a Shell Execute for arbitrary URIs.
             */
            IVsWebBrowsingService service = ServiceProvider.GetService(typeof(SVsWebBrowsingService)) as IVsWebBrowsingService;
            if (service != null)
            {
                __VSCREATEWEBBROWSER createFlags = __VSCREATEWEBBROWSER.VSCWB_AutoShow;
                VSPREVIEWRESOLUTION resolution = VSPREVIEWRESOLUTION.PR_Default;
                int result = ErrorHandler.CallWithCOMConvention(() => { ThreadHelper.ThrowIfNotOnUIThread(); return service.CreateExternalWebBrowser((uint)createFlags, resolution, uri.AbsoluteUri); });
                if (ErrorHandler.Succeeded(result))
                    return true;
            }

            // Fall back to Shell Execute, but only for http or https URIs
            if (uri.Scheme != "http" && uri.Scheme != "https")
                return false;

            try
            {
                Process.Start(uri.AbsoluteUri);
                return true;
            }
            catch (Win32Exception)
            {
            }
            catch (FileNotFoundException)
            {
            }

            return false;
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static Command1 Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this._package;
            }
        }

        public static string GetTaskItem(DTE2 dte)
        {
            var regex = new Regex(@"\W(\#[0-9]+\b)(?!;)", RegexOptions.Compiled | RegexOptions.IgnoreCase);

            // TODO: #2 -Get the current selected item from the task list
            ThreadHelper.ThrowIfNotOnUIThread();

            var items = dte.ToolWindows.TaskList.TaskItems;
            var itemList = items.Cast<TaskItem>();

            if (!itemList.Any())
            {
                return string.Empty;
            }

            MatchCollection matches = regex.Matches(itemList.FirstOrDefault<TaskItem>().Description);

            if (matches.Count == 0)
            {
                return string.Empty;
            }

            // When I debug, I can see what I need here... just don't seem to be able to access it in code, I must be doing something wrong :)
            //+Entry   { Microsoft.VisualStudio.Shell.TableControl.Implementation.SnapshotTableEntryViewModel}
            //Microsoft.VisualStudio.Shell.TableControl.ITableEntryHandle { Microsoft.VisualStudio.Shell.TableControl.Implementation.SnapshotTableEntryViewModel}

            return matches[0].Value.Substring(2);
        }

        public string GetGitUrl()
        {
            // TODO: #1 - Get the current repository's GIT URL from the GIT config file
            return "https://github.com/Vizioz/task-issues/issues/";
        }

        public bool HasGitUrl()
        {
            return string.IsNullOrEmpty(gitRepo) ? false : true;
        }
    }
}
