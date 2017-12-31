using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Xml;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.SlowCheetah;

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

        /// <summary>
        /// The transform file is the one the user selected.
        /// </summary>
        private string TransformFile { get; set; }

        /// <summary>
        /// The Destination file is either App or Web.config
        /// </summary>
        private string DestinationFile { get; set; }

        #endregion

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Debug.WriteLine("Instance Initialized");
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

        private void menuCommand_BeforeQueryStatus(object sender, EventArgs e)
        {
            Debug.WriteLine("Entering BeforeQueryStatus");
            // get the menu that fired the event
            var menuCommand = sender as OleMenuCommand;
            if (menuCommand != null)
            {
                Debug.WriteLine("Evaluating BeforeQueryStatus.");
                // start by assuming that the menu will not be shown
                menuCommand.Visible = false;
                menuCommand.Enabled = false;

                IVsHierarchy hierarchy = null;
                uint itemid = VSConstants.VSITEMID_NIL;

                if (!IsSingleProjectItemSelection(out hierarchy, out itemid)) return;
                Debug.WriteLine("Project Is NOT Selected.");

                var vsProject = (IVsProject)hierarchy;
                if (!ProjectSupportsTransforms(vsProject)) return;
                Debug.WriteLine("Project Supports Transforms.");

                if (!ItemSupportsTransforms(vsProject, itemid)) return;
                Debug.WriteLine("Selected Tranform and Destination Exist.");

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

            if (DestinationExists(itemFullPath)) { 
                this.TransformFile = itemFullPath;
            }

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

        private void ApplyTransform()
        {
            if(this.TransformFile == null ||
                this.DestinationFile == null)
            {
                Debug.WriteLine("Source Or Destination where null.  Nothing to do.");
                return;
            }
            ITransformer transformer = new XmlTransformer();
            transformer.Transform(this.DestinationFile, this.TransformFile, this.DestinationFile);

        }

        private bool DestinationExists(string sourceFile)
        {
            var transformFileInfo = new FileInfo(sourceFile);
            if(transformFileInfo.Name.StartsWith("web.", StringComparison.OrdinalIgnoreCase))
            {
               var tempDestination = sourceFile.Replace(transformFileInfo.Name, "Web.config");
                if (File.Exists(tempDestination))
                {
                    this.DestinationFile = tempDestination;
                    return true;
                }
            }
            if (transformFileInfo.Name.StartsWith("app.", StringComparison.OrdinalIgnoreCase))
            {
                this.DestinationFile = this.TransformFile.Replace(transformFileInfo.Name, "App.config");
                if (File.Exists(this.DestinationFile))
                {
                    return true;
                }
            }

            return false;
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
            Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.GetType().FullName));
            this.ApplyTransform();

        }
    }
}
