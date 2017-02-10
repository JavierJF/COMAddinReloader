using System;

//Needed ofr COM Interop
using System.Runtime.InteropServices;
using Application = Microsoft.Office.Interop.OneNote.Application;
using System.Windows.Forms;
using Extensibility;
using Microsoft.Office.Core;

namespace OneNoteAddIn
{
    [Guid("F432D536-D5AE-45C4-932E-B0F0AEF04C3D"), ProgId("OneNoteAddIn.LoadUnload")]
    public class LoadUnload : IDTExtensibility2, IRibbonExtensibility
    {
        COMAddIn lTools = null;
        COMAddIns addInsList = null;

        public string GetCustomUI(string RibbonID)
        {
            return Properties.Resources.ribbon;
        }

        public void reloadAddInLT(IRibbonControl control)
        {
            try
            {
                string lToolsId = lTools.ProgId;

                if (lTools.Connect)
                {
                    lTools.Connect = false;

                    MessageBox.Show("Wait until add-in reload: " + lToolsId + "\r\n");

                    addInsList.Update();
                }
                else
                {
                    lTools.Connect = true;
                    MessageBox.Show("Add-in: " + lTools.ToString() + " connected");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnBeginShutdown(ref Array custom)
        {
        }

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            Application onApp = (Application)Application;
            this.addInsList = onApp.COMAddIns;
            this.lTools = addInsList.Item(1);
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void OnStartupComplete(ref Array custom)
        {
        }
    }
}
