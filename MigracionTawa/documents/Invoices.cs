using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MigracionTawa.documents {
    class Invoices : MDocument {

        private const string INVOICES_SP_HEADER = "SEI_STW_PurchInvoices";
        private const string INVOICES_SP_LINES = "SEI_STW_PurchInvoicesLines ";
        private const string INVOICES_TABLE = "INT_Tawa_SBO.dbo.FacturasCabecera";
        private const string INVOICES_KEY = "IdFactura";

        public Invoices(SAPbobsCOM.Company Company)
            : base(Company, INVOICES_SP_HEADER, INVOICES_KEY) {
        }

        protected override void update(SAPbobsCOM.Company Company, bool successful, string id,string Code) {
            string updateString = "UPDATE " + INVOICES_TABLE + " SET INT_Estado = '";
            if (successful) {
                updateString += "P' ";
            } else {
                updateString += "E', INT_Error = '" + Company.GetLastErrorDescription().Replace('\'',' ') + "' ";
            }
            updateString += "WHERE IdFactura = " + id;
            updateRS.DoQuery(updateString);
        }

        protected override bool migrateDocuments(SAPbobsCOM.Company Company, SAPbobsCOM.Recordset migrationRS) {
            SAPbobsCOM.Documents invoice = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
            string sDocDate = migrationRS.Fields.Item("DocDate").Value;
            string sDueDate = migrationRS.Fields.Item("DocDueDate").Value;
            string sTaxDate = migrationRS.Fields.Item("TaxDate").Value;
            DateTime docDate = DateTime.ParseExact(sDocDate,"yyyyMMdd",System.Globalization.CultureInfo.InvariantCulture);
            DateTime dueDate = DateTime.ParseExact(sDueDate,"yyyyMMdd",System.Globalization.CultureInfo.InvariantCulture);
            DateTime taxDate = DateTime.ParseExact(sTaxDate,"yyyyMMdd",System.Globalization.CultureInfo.InvariantCulture);
            invoice.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service;
            invoice.CardCode = migrationRS.Fields.Item("CardCode").Value;            
            //invoice.ContactPersonCode = migrationRS.Fields.Item("CntPerson").Value;
            invoice.DocCurrency = migrationRS.Fields.Item("DocCurrency").Value;
            invoice.JournalMemo = migrationRS.Fields.Item("JournalMemo").Value;
            invoice.PaymentMethod = migrationRS.Fields.Item("MetodoPago").Value;
            invoice.NumAtCard = migrationRS.Fields.Item("NumAtCard").Value;
            invoice.ControlAccount = migrationRS.Fields.Item("ControlAccount").Value;
            invoice.Series = migrationRS.Fields.Item("Series").Value;
            invoice.UserFields.Fields.Item("U_MSS_REFACT_COT").Value = migrationRS.Fields.Item("Refacturable").Value;
            invoice.UserFields.Fields.Item("U_BPP_MDTD").Value = "01";
            //invoice.UserFields.Fields.Item("U_MSS_REFACT_COT").Value = migrationRS.Fields.Item("U_MSS_REFACT_COT").Value;
            invoice.DocDate = docDate;
            invoice.TaxDate = taxDate;
            invoice.DocDueDate = dueDate;

            SAPbobsCOM.Recordset lines = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            lines.DoQuery(INVOICES_SP_LINES + migrationRS.Fields.Item("IdFactura").Value);
            while (!lines.EoF) {
                invoice.Lines.AccountCode = lines.Fields.Item("AccountCode").Value;
                invoice.Lines.TaxCode = lines.Fields.Item("TaxCode").Value;
                invoice.Lines.LineTotal = lines.Fields.Item("LineTotal").Value;
                invoice.Lines.CostingCode = lines.Fields.Item("CostingCode").Value;
                invoice.Lines.CostingCode2 = lines.Fields.Item("CostingCode2").Value;
                invoice.Lines.CostingCode3 = lines.Fields.Item("CostingCode3").Value;
                invoice.Lines.CostingCode4 = lines.Fields.Item("CostingCode4").Value;
                invoice.Lines.CostingCode5 = lines.Fields.Item("CostingCode5").Value;
                invoice.Lines.ItemDescription = lines.Fields.Item("Description").Value;
                lines.MoveNext();
            }
            return invoice.Add() == 0;
        }
    }
}
