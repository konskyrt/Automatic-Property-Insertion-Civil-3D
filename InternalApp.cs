using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD;

using System.Windows.Forms;
using System.Reflection;

using Autodesk.Windows;

namespace AutoCadPlugIn
{

    public class InternalApplication : RibbonTool, IExtensionApplication
    {
        public void Initialize()
        {
            try
            {
                if (Autodesk.Windows.ComponentManager.Ribbon == null)
                {
                    Autodesk.Windows.ComponentManager.ItemInitialized += new EventHandler<RibbonItemEventArgs>(ComponentManager_ItemInitialized);
                }
                else
                {
                    AddAppEvent();
                    CreateFabTab();
                }
            }
            catch (System.Exception ex)
            {
                throw;
            }
        }


        private void ComponentManager_ItemInitialized(object sender, RibbonItemEventArgs e)
        {
            if (Autodesk.Windows.ComponentManager.Ribbon != null)
            {
                CreateFabTab();
                AddAppEvent();
                Autodesk.Windows.ComponentManager.ItemInitialized -= new EventHandler<RibbonItemEventArgs>(ComponentManager_ItemInitialized);
            }
        }

        private void AddAppEvent()
        {
            Autodesk.AutoCAD.ApplicationServices.Core.Application.SystemVariableChanged += new Autodesk.AutoCAD.ApplicationServices.SystemVariableChangedEventHandler(appSysVarChanged);
        }

        public void appSysVarChanged(object senderObj, Autodesk.AutoCAD.ApplicationServices.SystemVariableChangedEventArgs sysVarChEvtArgs)
        {
            if ((sysVarChEvtArgs.Name).ToString() == "WSCURRENT")
            {
                CreateFabTab();
            }
        }

        public void Terminate()
        {

        }
    }
}