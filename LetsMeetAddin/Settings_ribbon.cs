using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Settings_ribbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace LetsMeetAddin
{
    [ComVisible(true)]
    public class Settings_ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Settings_ribbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("LetsMeetAddin.Settings_ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226
        public string EditBoxGetTextLink(Office.IRibbonControl control)
        {
            String meetlink = Properties.Settings.Default.MeetLink;
            return meetlink;
        }

        public void editBoxLink_TextChanged(Office.IRibbonControl control, String text)
        {
            Properties.Settings.Default.MeetLink = text;
        }

        public string EditBoxGetTextName(Office.IRibbonControl control)
        {
            String meetlink = Properties.Settings.Default.MeetName;
            return meetlink;
        }

        public void editBoxName_TextChanged(Office.IRibbonControl control, String text)
        {
            Properties.Settings.Default.MeetName = text;
        }

        public string EditBoxGetTextDesc(Office.IRibbonControl control)
        {
            String meetlink = Properties.Settings.Default.MeetDesc;
            return meetlink;
        }

        public void editBoxDesc_TextChanged(Office.IRibbonControl control, String text)
        {
            Properties.Settings.Default.MeetDesc = text;
        }
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

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
