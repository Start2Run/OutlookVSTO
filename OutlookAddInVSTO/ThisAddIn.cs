using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools;
using WpfControlLibrary;

namespace OutlookAddInVSTO
{
    public partial class ThisAddIn
    {
        //https://github.com/officedev/outlook-add-in-command-demo
        private Outlook.Inspectors m_inspectors;
        private Outlook.AppointmentItem m_appointmentItem;
        private CustomTaskPane m_myCustomTaskPane;

        public Outlook.Application OutlookApplication;

        private Dictionary<Outlook.Inspector, InspectorWrapper> m_inspectorWrappersValue =
            new Dictionary<Outlook.Inspector, InspectorWrapper>();

        public Dictionary<Outlook.Inspector, InspectorWrapper> InspectorWrappers => m_inspectorWrappersValue;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {

            OutlookApplication = this.Application;
            m_inspectors = this.Application.Inspectors;
            m_inspectors.NewInspector += Inspectors_NewInspector;

            foreach (Outlook.Inspector inspector in m_inspectors)
            {
                Inspectors_NewInspector(inspector);
            }

            OutlookApplication.ItemSend += OutlookApplication_ItemSend;

            var myUserControl1 = new UserControl();
            m_myCustomTaskPane = this.CustomTaskPanes.Add(myUserControl1, "My Task Pane");
            m_myCustomTaskPane.Visible = true;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785

            m_inspectors.NewInspector -= Inspectors_NewInspector;
            OutlookApplication.ItemSend -= OutlookApplication_ItemSend;
            m_inspectors = null;
            m_inspectorWrappersValue = null;
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

        void Inspectors_NewInspector(Outlook.Inspector inspector)
        {
            if (inspector.CurrentItem is Outlook.AppointmentItem appointmentItem)
            {
                m_appointmentItem = appointmentItem;
                m_inspectorWrappersValue.Add(inspector, new InspectorWrapper(inspector));
            }
        }

        void OutlookApplication_ItemSend(object item, ref bool cancel)
        {
            const string prSmtpAddress = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
            var recipients = m_appointmentItem.Recipients;
            foreach (Outlook.Recipient recipient in recipients)
            {
                var pa = recipient.PropertyAccessor;
                string smtpAddress = pa.GetProperty(prSmtpAddress).ToString();
                Debug.WriteLine(recipient.Name + " SMTP=" + smtpAddress);
            }
        }
    }

    public class InspectorWrapper
    {
        private Outlook.Inspector m_inspector;
        private CustomTaskPane m_taskPane;
        public InspectorWrapper(Outlook.Inspector Inspector)
        {
            m_inspector = Inspector;
            ((Outlook.InspectorEvents_Event)m_inspector).Close +=
                InspectorWrapper_Close;

            var host = new ElementHost {Dock = DockStyle.Fill, Child = new UserControlRibbon()};

            var uc = new UserControl();
            uc.Controls.Add(host);
            m_taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(uc, "My task pane", m_inspector);
            m_taskPane.VisibleChanged += TaskPane_VisibleChanged;
        }

        void TaskPane_VisibleChanged(object sender, EventArgs e)
        {
            Globals.Ribbons[m_inspector].RibbonMain.ShowTaskPaneBtn.Checked =
                m_taskPane.Visible;
        }
        void InspectorWrapper_Close()
        {
            if (m_taskPane != null)
            {
                Globals.ThisAddIn.CustomTaskPanes.Remove(m_taskPane);
            }

            m_taskPane = null;
            Globals.ThisAddIn.InspectorWrappers.Remove(m_inspector);
            ((Outlook.InspectorEvents_Event)m_inspector).Close -=
                InspectorWrapper_Close;
            m_inspector = null;
        }

        public CustomTaskPane CustomTaskPane => m_taskPane;
    }
}
