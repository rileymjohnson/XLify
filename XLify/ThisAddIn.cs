using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Diagnostics;
using System.IO;

namespace XLify
{
    public partial class ThisAddIn
    {
        private CustomTaskPane _taskPane;
        private MyTaskPaneControl _taskPaneControl;
        private const int BasePaneWidthDips = 360; // logical (96-DPI) pixels

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                // Route app stdout/stderr/Trace to Seq early
                try { LoggingBridge.EnableSeqForApp("http://localhost:5341"); } catch { }

                this.Application.StatusBar = "XLify add-in loaded (debug)";
                this.Application.WindowResize += new Excel.AppEvents_WindowResizeEventHandler(Application_WindowResize);
                this.Application.WindowActivate += new Excel.AppEvents_WindowActivateEventHandler(Application_WindowActivate);

            }
            catch { }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            try
            {
                var app = this.Application;
                if (app != null)
                {
                    app.StatusBar = "XLify ribbon loading (debug)";
                }
            }
            catch { }
            return new XLifyRibbon();
        }

        internal void ShowTaskPane()
        {
            try
            {
                if (_taskPane == null)
                {
                    if (this.CustomTaskPanes == null)
                    {
                        // If the collection hasn't been initialized yet, bail gracefully.
                        return;
                    }
                    _taskPaneControl = new MyTaskPaneControl();
                    _taskPane = this.CustomTaskPanes.Add(_taskPaneControl, "XLify");
                    _taskPane.VisibleChanged += TaskPane_VisibleChanged;
                    try
                    {
                        _taskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                        _taskPane.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal;
                    }
                    catch { }
                    _taskPane.Width = 360;
                }
                UpdateTaskPaneWidth();
                _taskPane.Visible = true;
            }
            catch { }
        }

        internal void ToggleTaskPane()
        {
            if (_taskPane == null)
            {
                ShowTaskPane();
            }
            else
            {
                _taskPane.Visible = !_taskPane.Visible;
            }
        }

        internal void SetTaskPaneVisible(bool visible)
        {
            try
            {
                if (visible)
                {
                    ShowTaskPane();
                }
                else if (_taskPane != null)
                {
                    _taskPane.Visible = false;
                }
            }
            catch { }
        }

        internal bool IsTaskPaneVisible()
        {
            try { return _taskPane != null && _taskPane.Visible; } catch { return false; }
        }

        private void TaskPane_VisibleChanged(object sender, EventArgs e)
        {
            XLifyRibbon.Invalidate("btnOpenTaskPane");
            if (_taskPane != null && _taskPane.Visible)
            {
                UpdateTaskPaneWidth();
            }
        }

        private void Application_WindowResize(Excel.Workbook Wb, Excel.Window Wn)
        {
            UpdateTaskPaneWidth();
        }

        private void Application_WindowActivate(Excel.Workbook Wb, Excel.Window Wn)
        {
            UpdateTaskPaneWidth();
        }

        private void UpdateTaskPaneWidth()
        {
            try
            {
                if (_taskPane == null) return;
                IntPtr hwnd = IntPtr.Zero;
                try { hwnd = new IntPtr(this.Application?.Hwnd ?? 0); } catch { }
                double scale = DpiHelper.GetScaleForWindow(hwnd);
                int target = (int)Math.Round(BasePaneWidthDips * scale);
                // Clamp to reasonable bounds
                if (target < 240) target = 240;
                if (target > 800) target = 800;
                _taskPane.Width = target;
            }
            catch { }
        }
    }
}
