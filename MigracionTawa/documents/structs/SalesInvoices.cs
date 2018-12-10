using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MigracionTawa.documents.structs {
    class SalesInvoices : MDocument {

        private const string INVOICES_HEADER_SP = "SEI_STW_PayrollInvoices";
        private const string INVOICES_TABLE = "INT_Tawa_SBO.dbo.FacturasPayroll";
        private const string INVOICES_KEY = "IdFactura";

        public SalesInvoices(SAPbobsCOM.Company Company)
            : base(Company, INVOICES_HEADER_SP, INVOICES_KEY) {
        }

        protected override void update(SAPbobsCOM.Company Company, bool successful, string id, string Code) {
            string updateString = "UPDATE " + INVOICES_TABLE + " SET INT_Estado = '";
            if (successful) {
                updateString += "P' ";
            } else {
                updateString += "E', INT_Error = '" + Company.GetLastErrorDescription().Replace('\'', ' ') + "' ";
            }
            updateString += "WHERE IdFactura = " + id;
            updateRS.DoQuery(updateString);
        }

        protected override bool migrateDocuments(SAPbobsCOM.Company Company, SAPbobsCOM.Recordset migrationRS) {
            SAPbobsCOM.Documents salesInvoice = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
            salesInvoice.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices;
            salesInvoice.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service;
            string invDate = migrationRS.Fields.Item("INT_Fecha").Value;
            DateTime pDate = DateTime.ParseExact(invDate, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
            salesInvoice.CardCode = migrationRS.Fields.Item("CardCode").Value;
            salesInvoice.DocDate = pDate;
            salesInvoice.TaxDate = pDate;
            salesInvoice.DocCurrency = migrationRS.Fields.Item("DocCurrency").Value;
            salesInvoice.ControlAccount = migrationRS.Fields.Item("AccountCode").Value;
            salesInvoice.GroupNumber = migrationRS.Fields.Item("CondicionPago").Value;
            salesInvoice.UserFields.Fields.Item("U_MSS_GRPFACT").Value = migrationRS.Fields.Item("CO_GRUP_FACT").Value;
            //salesInvoice.UserFields.Fields.Item("U_MSS_FECREC").Value = pDate;

            salesInvoice.Lines.LineTotal = migrationRS.Fields.Item("IM_BRUT_FACT").Value;
            salesInvoice.Lines.TaxCode = migrationRS.Fields.Item("IM_BRUT_IGV").Value > 0 ? "IGV" : "IGV_EXE";
            salesInvoice.Lines.AccountCode = migrationRS.Fields.Item("AccountCodeDet").Value;
            salesInvoice.Lines.ItemDescription = migrationRS.Fields.Item("NO_GLOS_SAP1").Value;
            salesInvoice.Lines.CostingCode = migrationRS.Fields.Item("CostingCode").Value;
            salesInvoice.Lines.CostingCode2 = migrationRS.Fields.Item("CostingCode2").Value;
            salesInvoice.Lines.CostingCode3 = migrationRS.Fields.Item("CostingCode3").Value;
            salesInvoice.Lines.CostingCode4 = migrationRS.Fields.Item("CostingCode4").Value;
            salesInvoice.Lines.CostingCode5 = migrationRS.Fields.Item("CostingCode5").Value;

            return salesInvoice.Add() == 0;
        }
    }
}
