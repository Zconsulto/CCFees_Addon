using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CCFees_Addon
{
    [FormAttribute("133", "SystemForm1.b1f")]
    class SystemForm1 : SystemFormBase
    {
        public SystemForm1()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_0").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_1").Specific));
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.CalculateButton);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_2").Specific));
            this.Button1.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.AddButton);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.EditText EditText0;

        private void OnCustomInitialize()
        {

        }

        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;

        public SAPbouiCOM.EditText getEdittext(SAPbouiCOM.Form Form, int itemCode)
        {
            return ((SAPbouiCOM.EditText)Form.Items.Item(itemCode.ToString()).Specific);

        }
        public string getEdittextString(SAPbouiCOM.Form Form, int itemCode)
        {
            return ((SAPbouiCOM.EditText)Form.Items.Item(itemCode.ToString()).Specific).Value;

        }
        public SAPbouiCOM.Matrix getMatrix(SAPbouiCOM.Form activeForm, int itemCode)
        {
            return (SAPbouiCOM.Matrix)activeForm.Items.Item(itemCode.ToString()).Specific;

        }
        public string getMatirxString(SAPbouiCOM.Matrix matrix, int Column, int row)
        {
            return (((SAPbouiCOM.EditText)matrix.Columns.Item(Column.ToString()).Cells.Item(row).Specific).Value.ToString());

        }
        public SAPbouiCOM.ComboBox getComboBox(SAPbouiCOM.Form activeForm, int itemCode)
        {
            return (SAPbouiCOM.ComboBox)activeForm.Items.Item(itemCode.ToString()).Specific;

        }
        public int printMessageBox(String message, int buttons, string[] button_titles)
        {
            if (button_titles.Length == 1)
            {
                return Application.SBO_Application.MessageBox(message, buttons, button_titles[0]);
            }
            else if (button_titles.Length == 2)
            {
                return Application.SBO_Application.MessageBox(message, buttons, button_titles[0], button_titles[1]);
            }
            else if (button_titles.Length == 3)
            {
                return Application.SBO_Application.MessageBox(message, buttons, button_titles[0], button_titles[1], button_titles[2]);
            }
            else
            {
                return Application.SBO_Application.MessageBox(message, buttons);
            }

        }
        public SAPbobsCOM.Recordset query()
        {
            SAPbobsCOM.Company oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
            return (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        }
        public string getQueryField(SAPbobsCOM.Recordset recordSet, string field)
        {
            return recordSet.Fields.Item(field).Value.ToString();
        }
        /*
         * Message Times=[ 1-->Short || 2-->Medium || 3-->Long ]
         * Message Type=[ 0-->Error || 1-->None || 2-->success || 3-->Warning ]
         * */
        public void printStatusBar(String message, int messageTime, int MessageType)
        {
            SAPbouiCOM.BoMessageTime messageTime_bmt = SAPbouiCOM.BoMessageTime.bmt_Short;
            if (messageTime == 2)
            {
                messageTime_bmt = SAPbouiCOM.BoMessageTime.bmt_Medium;
            }
            else if (messageTime == 3)
            {
                messageTime_bmt = SAPbouiCOM.BoMessageTime.bmt_Long;
            }


            //-------
            SAPbouiCOM.BoStatusBarMessageType messageType_smt = SAPbouiCOM.BoStatusBarMessageType.smt_None;

            if (MessageType == 0)
            {
                messageType_smt = SAPbouiCOM.BoStatusBarMessageType.smt_Error;
            }
            else if (MessageType == 2)
            {
                messageType_smt = SAPbouiCOM.BoStatusBarMessageType.smt_Success;
            }
            else if (MessageType == 3)
            {
                messageType_smt = SAPbouiCOM.BoStatusBarMessageType.smt_Warning;
            }



            Application.SBO_Application.StatusBar.SetText(message, messageTime_bmt, messageType_smt);

        }
        public void Click(SAPbouiCOM.Form activeForm, int itemCode)
        {
            activeForm.Items.Item(itemCode.ToString()).Click();
        }
        private void setMatrixValue(SAPbouiCOM.Matrix Matrix, int column, int row, string value)
        {
            ((SAPbouiCOM.EditText)Matrix.Columns.Item(column.ToString()).Cells.Item(row).Specific).Value = value;
        }
        private bool created(SAPbouiCOM.Form activeForm)
        {
            string docNumStr = getEdittextString(activeForm, 8);
            int docNum = int.Parse(docNumStr);
            SAPbobsCOM.Recordset invoices = query();
            invoices.DoQuery("SELECT T0.\"DocNum\" FROM OINV T0 WHERE T0.\"DocNum\" = '" + docNum + "'");
            string docNumReturn = getQueryField(invoices, "DocNum");
            return !docNumReturn.Equals("0");
        }
        private bool noTaxRows(int rowCount, SAPbouiCOM.Matrix matrix)
        {
            bool EmptyRowTax = false;
            for (int i = 1; i <= rowCount - 1; i++)
            {
                string taxCode = getMatirxString(matrix, 160, i);

                if (string.IsNullOrEmpty(taxCode))
                {
                    EmptyRowTax = true;
                    break;
                }
            }
            return EmptyRowTax;

        }
        private bool noRows(SAPbouiCOM.Form activeForm)
        {
            SAPbouiCOM.Matrix matrix = getMatrix(activeForm, 38); ;
            int rowCount = matrix.VisualRowCount;
            return (rowCount == 1);
        }
        private void CalculateButton(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            printMessageBox("Calculate Button Working", 0, new string[] { "Yes", "No" });

        }

        private void AddButton(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            printMessageBox("Add Button Working", 0, new string[] { "Yes", "No" });

        }
    }
}
