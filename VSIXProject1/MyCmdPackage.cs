using EnvDTE;
using EnvDTE80;
using ItemColorPackage;
using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel.Design;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.InteropServices;

namespace VSIXProject1
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0")]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(MyCmdPackage.PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    public sealed class MyCmdPackage : Package
    {
        public const string PackageGuidString = "7a9f8e8d-ef24-4c98-b859-fc0d4c59d5a9";
        #region Package Members
        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            //MyCmd.Initialize(this);
            base.Initialize();


            OleMenuCommandService oOleMenuCommandService = (OleMenuCommandService)GetService(typeof(IMenuCommandService));

            if (oOleMenuCommandService != null)
            {
                CommandID oCommandID = new CommandID(Consts.ItemColorPackage, (int)Consts.Command);
                CommandID oCommandListID = new CommandID(Consts.ItemColorPackage, (int)Consts.CommandList);

                OleMenuCommand oCommand = new OleMenuCommand(CommandInvoke, oCommandID);
                OleMenuCommand oCommandList = new OleMenuCommand(CommandListInvoke, oCommandListID);

                oCommand.BeforeQueryStatus += CommandQueryStatus;

                oOleMenuCommandService.AddCommand(oCommand);
                oOleMenuCommandService.AddCommand(oCommandList);
            }
        }

        private void CommandInvoke(object sender, EventArgs e)
        {
            OleMenuCmdEventArgs oOleMenuCmdEventArgs = (OleMenuCmdEventArgs)e;

            DTE2 Application = (DTE2)GetService(typeof(DTE));
            ProjectItem oProjectItem = Application.SelectedItems.Item(1).ProjectItem;
            IServiceProvider oServiceProvider = new ServiceProvider((Microsoft.VisualStudio.OLE.Interop.IServiceProvider)Application);
            Microsoft.Build.Evaluation.Project oBuildProject = Microsoft.Build.Evaluation.ProjectCollection.GlobalProjectCollection.GetLoadedProjects(oProjectItem.ContainingProject.FullName).SingleOrDefault();
            Microsoft.Build.Evaluation.ProjectProperty oGUID = oBuildProject.AllEvaluatedProperties.SingleOrDefault(oProperty => oProperty.Name == "ProjectGuid");
            Microsoft.VisualStudio.Shell.Interop.IVsHierarchy oVsHierarchy = VsShellUtilities.GetHierarchy(oServiceProvider, new Guid(oGUID.EvaluatedValue));
            Microsoft.VisualStudio.Shell.Interop.IVsBuildPropertyStorage oVsBuildPropertyStorage = (Microsoft.VisualStudio.Shell.Interop.IVsBuildPropertyStorage)oVsHierarchy;

            string szItemPath = (string)oProjectItem.Properties.Item("FullPath").Value;
            uint nItemId;

            oVsHierarchy.ParseCanonicalName(szItemPath, out nItemId);

            if (oOleMenuCmdEventArgs.OutValue != IntPtr.Zero)
            {
                string szOut;

                oVsBuildPropertyStorage.GetItemAttribute(nItemId, "ItemColor", out szOut);

                Marshal.GetNativeVariantForObject(szOut, oOleMenuCmdEventArgs.OutValue);
            }
            else if (oOleMenuCmdEventArgs.InValue != null)
            {
                oVsBuildPropertyStorage.SetItemAttribute(nItemId, "ItemColor", Convert.ToString(oOleMenuCmdEventArgs.InValue));
            }
        }

        private void CommandQueryStatus(object sender, EventArgs e)
        {
            OleMenuCommand oCommand = (OleMenuCommand)sender;

            oCommand.Visible = ItemToHandle();
        }

        private void CommandListInvoke(object sender, EventArgs e)
        {
            OleMenuCmdEventArgs oOleMenuCmdEventArgs = (OleMenuCmdEventArgs)e;

            if (oOleMenuCmdEventArgs.OutValue != IntPtr.Zero)
            {
                Marshal.GetNativeVariantForObject(new string[] { "Red", "Green", "Blue" }, oOleMenuCmdEventArgs.OutValue);
            }
        }

        private bool ItemToHandle()
        {
            DTE2 Application = (DTE2)GetService(typeof(DTE));

            return ((Application.SelectedItems.Count == 1) && (Application.SelectedItems.Item(1).ProjectItem != null) && (Application.SelectedItems.Item(1).ProjectItem.Name.Contains(".cs")));
        }
        #endregion
    }
}
