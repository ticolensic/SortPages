using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices.ComTypes;
using System.Xml.Linq;

using Extensibility;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.OneNote;
using App = Microsoft.Office.Interop.OneNote.Application;
using HS = Microsoft.Office.Interop.OneNote.HierarchyScope;
using System.Diagnostics.CodeAnalysis;



namespace SortPages
{
    [Guid("573DE9A8-7ED9-446C-8B27-1BFFDC78D150")]
    [ProgId("SortPages.Connect")]

    public class Connect : Object, Extensibility.IDTExtensibility2, IRibbonExtensibility
    {
        private App oneNote;

        public void AddInButtonClicked(IRibbonControl control)
        {
            string sectionXml;
            
            oneNote.GetHierarchy(
                oneNote.Windows.CurrentWindow.CurrentSectionId,
                HS.hsPages, 
                out sectionXml,
                Microsoft.Office.Interop.OneNote.XMLSchema.xs2010);

            var sectionNode = XDocument.Parse(sectionXml);
            var ns = sectionNode.Root.Name.Namespace;

            switch (control.Id)
            {
                case "DMdesc":
                    Tools.SortByAttribute(sectionNode.Element(ns + "Section"), "dateTime", false);
                    break;
                case "DMasc":
                    Tools.SortByAttribute(sectionNode.Element(ns + "Section"), "dateTime");
                    break;
                case "LMdesc":
                    Tools.SortByAttribute(sectionNode.Element(ns + "Section"), "lastModifiedTime", false);
                    break;
                case "LMasc":
                    Tools.SortByAttribute(sectionNode.Element(ns + "Section"), "lastModifiedTime");
                    break;
                default:
                    throw new System.ArgumentException("Wrong button id"); 
            }
            oneNote.UpdateHierarchy(sectionNode.ToString(),
                Microsoft.Office.Interop.OneNote.XMLSchema.xs2010);
        
        }



        // implement IRibbonExtensibility
        public string GetCustomUI(string RibbonID)
        {
            return Properties.Resources.ribbon;
        }
    
        // implement IDTExtensibility2: OnConnection, OnAddInsUpdate, OnBeginShutdown, OnDisconnection
        [STAThread]
        public void OnConnection (object application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            oneNote = (App)application;
        }

        public void OnAddInsUpdate(ref Array custom)
        { }
        
        public void OnBeginShutdown(ref Array custom)
        { }

        [SuppressMessage("Microsoft.Reliability", "CA2001:AvoidCallingProblematicMethods", MessageId = "System.GC.Collect")]
        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            oneNote = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect(); // maybe we need this, maybe not, who knows?
        }

        public void OnStartupComplete(ref Array custom)
        { }
        
    }

}
