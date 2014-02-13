using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Res = Tabbles.OutlookAddIn.Properties.Resources;
using System.Drawing;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace Tabbles.OutlookAddIn
{
    [ComVisible(true)]
    public class TabblesRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        //public event EventHandler TagUsingTabbles;
        //public event EventHandler OpenInTabbles;
        //public event EventHandler TabblesSearch;
        //public event EventHandler SyncWithTabbles;

        //public event IsAnyEmailSelectedHandler IsAnyEmailSelected;

        public MenuManager mMenuManager;
        public TabblesRibbon(MenuManager menuM)
        {
            mMenuManager = menuM;
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            if (ribbonID == "Microsoft.Outlook.Explorer")
            {
                return GetResourceText("Tabbles.OutlookAddIn.RibbonExplorer.xml");
            }
            else if (ribbonID == "Microsoft.Outlook.Mail.Compose" ||
                ribbonID == "Microsoft.Outlook.Mail.Read")
            {
                return GetResourceText("Tabbles.OutlookAddIn.RibbonInspector.xml");
            }

            return null;
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, select the Ribbon XML item in Solution Explorer and then press F1

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnAction(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "tagUsingTabblesButton":
                case "tagUsingTabblesMenuSingle":
                case "tagUsingTabblesMenuMultiple":
                    mMenuManager.TagSelectedEmailsWithTabbles();
                    break;
                case "openInTabblesButton":
                case "openInTabblesMenu":
                    //if (OpenInTabbles != null)
                    //{
                    //    OpenInTabbles(control, EventArgs.Empty);
                    //}
                    break;
                case "tabblesSearchButton":
                    //if (TabblesSearch != null)
                    //{
                    //    TabblesSearch(control, EventArgs.Empty);
                    //}
                    break;
                case "syncWithTabblesButton":
                    //if (SyncWithTabbles != null)
                    //{
                    //    SyncWithTabbles(control, EventArgs.Empty);
                    //}
                    break;
                default:
                    break;
            }
        }

        public string GetLabel(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "tagUsingTabblesButton":
                case "tagUsingTabblesMenuSingle":
                case "tagUsingTabblesMenuMultiple":
                    return Res.MenuTagUsingTabbles;
                case "openInTabblesButton":
                case "openInTabblesMenu":
                    return Res.MenuOpenInTabbles;
                case "tabblesSearchButton":
                    return Res.MenuTabblesSearch;
                case "syncWithTabblesButton":
                    return Res.MenuSyncWithTabbles;
                default:
                    return string.Empty;
            }
        }

        public Bitmap OnLoadImage(string imageName)
        {
            Bitmap image = null;

            switch (imageName)
            {
                case "tag_using_tabbles":
                    image = Res.tag_using_tabbles_32;
                    break;
                case "open_in_tabbles":
                    image = Res.open_in_tabbles_32;
                    break;
                case "search":
                    image = Res.search_32;
                    break;
                case "sync_with_tabbles":
                    image = Res.sync_with_tabbles_32;
                    break;
                case "tag_using_tabbles_small":
                    image = Res.tag_using_tabbles_16_png;
                    break;
                case "open_in_tabbles_small":
                    image = Res.open_in_tabbles_16_png;
                    break;
                default:
                    break;
            }

            return image;
        }

        //public bool IsAnythingSelected(Office.IRibbonControl control)
        //{
        //    return IsAnyEmailSelected != null && IsAnyEmailSelected();
        //}

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
