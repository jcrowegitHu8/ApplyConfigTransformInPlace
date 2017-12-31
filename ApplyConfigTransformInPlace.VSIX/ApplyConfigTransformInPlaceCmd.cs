using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Xml;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace ApplyConfigTransformInPlace.VSIX
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class ApplyConfigTransformInPlaceCmd
    {

        #region Variables & Properties
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("97f88669-0014-472e-943c-b4166c83bb30");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static ApplyConfigTransformInPlaceCmd Instance
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
                return this.package;
            }
        }

        #endregion

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new ApplyConfigTransformInPlaceCmd(package);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ApplyConfigTransformInPlaceCmd"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private ApplyConfigTransformInPlaceCmd(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                Console.WriteLine("Command Evaluating");
                var menuCommandID = new CommandID(CommandSet, CommandId);

                // WE COMMENT OUT THE LINE BELOW
                //var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);

                // AND REPLACE IT WITH A DIFFERENT TYPE
                var menuItem = new OleMenuCommand(MenuItemCallback, menuCommandID);
                menuItem.BeforeQueryStatus += menuCommand_BeforeQueryStatus;

                commandService.AddCommand(menuItem);
            }
        }

        private void menuCommand_BeforeQueryStatus_Old(object sender, EventArgs e)
        {
            // get the menu that fired the event
            var menuCommand = sender as OleMenuCommand;
            if (menuCommand != null)
            {
                // start by assuming that the menu will not be shown
                menuCommand.Visible = false;
                menuCommand.Enabled = false;

                IVsHierarchy hierarchy = null;
                uint itemid = VSConstants.VSITEMID_NIL;

                if (!IsSingleProjectItemSelection(out hierarchy, out itemid)) return;
                // Get the file path
                string itemFullPath = null;
                ((IVsProject)hierarchy).GetMkDocument(itemid, out itemFullPath);
                var transformFileInfo = new FileInfo(itemFullPath);

                // then check if the file is named 'web.config'
                bool isConfig = transformFileInfo.Name.EndsWith(".config", StringComparison.OrdinalIgnoreCase);

                // if not leave the menu hidden
                if (!isConfig) return;

                menuCommand.Visible = true;
                menuCommand.Enabled = true;
            }
        }

        private void OnBeforeQueryStatusAddTransformCommand(object sender, EventArgs e)
        {
            // get the menu that fired the event
            var menuCommand = sender as OleMenuCommand;
            if (menuCommand != null)
            {
                // start by assuming that the menu will not be shown
                menuCommand.Visible = false;
                menuCommand.Enabled = false;

                IVsHierarchy hierarchy = null;
                uint itemid = VSConstants.VSITEMID_NIL;

                if (!IsSingleProjectItemSelection(out hierarchy, out itemid)) return;

                var vsProject = (IVsProject)hierarchy;
                if (!ProjectSupportsTransforms(vsProject)) return;

                if (!ItemSupportsTransforms(vsProject, itemid)) return;

                menuCommand.Visible = true;
                menuCommand.Enabled = true;
            }
        }

        static List<string> SupportedProjectExtensions = new List<string>
        {
            ".csproj",
            ".vbproj",
            ".fsproj"
        };

        private bool ProjectSupportsTransforms(IVsProject project)
        {
            string projectFullPath = null;
            if (ErrorHandler.Failed(project.GetMkDocument(VSConstants.VSITEMID_ROOT, out projectFullPath))) return false;

            string projectExtension = Path.GetExtension(projectFullPath);

            foreach (string supportedExtension in SupportedProjectExtensions)
            {
                if (projectExtension.Equals(supportedExtension, StringComparison.InvariantCultureIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }


        private bool ItemSupportsTransforms(IVsProject project, uint itemid)
        {
            string itemFullPath = null;

            if (ErrorHandler.Failed(project.GetMkDocument(itemid, out itemFullPath))) return false;

            //TODO 

            var transformFileInfo = new FileInfo(itemFullPath);
            bool isConfig = transformFileInfo.Name.EndsWith(".config", StringComparison.OrdinalIgnoreCase);

            return (isConfig && IsXmlFile(itemFullPath));
        }


        private bool IsXmlFile(string filepath)
        {
            if (string.IsNullOrWhiteSpace(filepath)) { throw new ArgumentNullException("filepath"); }
            if (!File.Exists(filepath)) throw new FileNotFoundException("File not found", filepath);

            var isXmlFile = true;
            try
            {
                using (var xmlTextReader = new XmlTextReader(filepath))
                {
                    // This is required because if the XML file has a DTD then it will try and download the DTD!
                    xmlTextReader.DtdProcessing = DtdProcessing.Ignore;
                    xmlTextReader.Read();
                }
            }
            catch (XmlException)
            {
                isXmlFile = false;
            }
            return isXmlFile;
        }

        public static bool IsSingleProjectItemSelection(out IVsHierarchy hierarchy, out uint itemid)
        {
            hierarchy = null;
            itemid = VSConstants.VSITEMID_NIL;
            int hr = VSConstants.S_OK;

            var monitorSelection = Package.GetGlobalService(typeof(SVsShellMonitorSelection)) as IVsMonitorSelection;
            var solution = Package.GetGlobalService(typeof(SVsSolution)) as IVsSolution;
            if (monitorSelection == null || solution == null)
            {
                return false;
            }

            IVsMultiItemSelect multiItemSelect = null;
            IntPtr hierarchyPtr = IntPtr.Zero;
            IntPtr selectionContainerPtr = IntPtr.Zero;

            try
            {
                hr = monitorSelection.GetCurrentSelection(out hierarchyPtr, out itemid, out multiItemSelect, out selectionContainerPtr);

                if (ErrorHandler.Failed(hr) || hierarchyPtr == IntPtr.Zero || itemid == VSConstants.VSITEMID_NIL)
                {
                    // there is no selection
                    return false;
                }

                // multiple items are selected
                if (multiItemSelect != null) return false;

                // there is a hierarchy root node selected, thus it is not a single item inside a project

                if (itemid == VSConstants.VSITEMID_ROOT) return false;

                hierarchy = Marshal.GetObjectForIUnknown(hierarchyPtr) as IVsHierarchy;
                if (hierarchy == null) return false;

                Guid guidProjectID = Guid.Empty;

                if (ErrorHandler.Failed(solution.GetGuidOfProject(hierarchy, out guidProjectID)))
                {
                    return false; // hierarchy is not a project inside the Solution if it does not have a ProjectID Guid
                }

                // if we got this far then there is a single project item selected
                return true;
            }
            finally
            {
                if (selectionContainerPtr != IntPtr.Zero)
                {
                    Marshal.Release(selectionContainerPtr);
                }

                if (hierarchyPtr != IntPtr.Zero)
                {
                    Marshal.Release(hierarchyPtr);
                }
            }
        }



        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            string message = string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.GetType().FullName);
            string title = "Apply Config Transform In Place";

            // Show a message box to prove we were here
            VsShellUtilities.ShowMessageBox(
                this.ServiceProvider,
                message,
                title,
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }
    }
}
