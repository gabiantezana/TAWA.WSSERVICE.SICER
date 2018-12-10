using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MigracionTawa.documents {
    class WebInvoices {
        private const string INVOICES_SP_HEADER = "EXEC [SEI_STW_FacturacionWeb]";
        private const string INVOICES_SP_LINES = "EXEC [SEI_STW_FacturacionWebDetalle] ";
        private const string INVOICES_TABLE = "INT_Tawa_SBO.dbo.FacturasWeb";

        public void migrate(SAPbobsCOM.Company Company) {
            SAPbobsCOM.Recordset migrationRS = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset updateRS = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            migrationRS.DoQuery(INVOICES_SP_HEADER);
            while (!migrationRS.EoF) {
                migrateDocuments(Company, migrationRS);
                migrationRS.MoveNext();
            }
        }

        protected void migrateDocuments(SAPbobsCOM.Company Company, SAPbobsCOM.Recordset migrationRS) {
            int exCode = migrationRS.Fields.Item("ExCode").Value;
            int tipoDoc = migrationRS.Fields.Item("TipoDocumento").Value;
            int etapa = migrationRS.Fields.Item("Etapa").Value;
            int idFactura = migrationRS.Fields.Item("IdFactura").Value;
            SAPbobsCOM.Recordset updateRS = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try {
                SAPbobsCOM.Documents invoice;
                int docSubType = migrationRS.Fields.Item("DocSubType").Value;
                switch (docSubType) {
                    case 18:
                        invoice = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                        break;
                    case 19:
                        invoice = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes);
                        break;
                    case 181:
                        invoice = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                        break;
                    default:
                        invoice = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                        break;
                }

                invoice.UserFields.Fields.Item("U_BPP_MDTD").Value = migrationRS.Fields.Item("U_BPP_MDTD").Value;
                string sDocDate = migrationRS.Fields.Item("DocDate").Value;
                string sDueDate = migrationRS.Fields.Item("DocDueDate").Value;
                string sTaxDate = migrationRS.Fields.Item("TaxDate").Value;
                DateTime docDate = DateTime.ParseExact(sDocDate, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                DateTime dueDate = DateTime.ParseExact(sDueDate, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                DateTime taxDate = DateTime.ParseExact(sTaxDate, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                invoice.Series = migrationRS.Fields.Item("Series").Value;
                invoice.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service;
                invoice.CardCode = migrationRS.Fields.Item("CardCode").Value;
                invoice.DocCurrency = migrationRS.Fields.Item("DocCurrency").Value;
                invoice.JournalMemo = migrationRS.Fields.Item("JournalMemo").Value;
                invoice.Comments = migrationRS.Fields.Item("Asunto").Value;
                invoice.PaymentMethod = migrationRS.Fields.Item("MetodoPago").Value;
                invoice.NumAtCard = migrationRS.Fields.Item("NumAtCard").Value;
                invoice.ControlAccount = migrationRS.Fields.Item("ControlAccount").Value;
                invoice.FolioPrefixString = migrationRS.Fields.Item("FolioPref").Value;
                invoice.FolioNumber = migrationRS.Fields.Item("FolioNum").Value;
                invoice.UserFields.Fields.Item("U_MSS_REFACT_COT").Value = migrationRS.Fields.Item("Refacturable").Value;
                invoice.UserFields.Fields.Item("U_MSS_Refacturado").Value = "Y";
                invoice.DocDate = docDate;
                invoice.TaxDate = taxDate;
                invoice.DocDueDate = dueDate;
                invoice.UserFields.Fields.Item("U_ExCode").Value = exCode;
                invoice.UserFields.Fields.Item("U_Etapa").Value = etapa;
                invoice.UserFields.Fields.Item("U_WebType").Value = tipoDoc;

                SAPbobsCOM.Recordset lines = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                lines.DoQuery(INVOICES_SP_LINES + "'" + exCode + "', '" + tipoDoc + "', '" + etapa + "', '" + migrationRS.Fields.Item("IdFactura").Value + "'");
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
                    invoice.Lines.Add();
                    lines.MoveNext();
                }
                Company.StartTransaction();
                if (invoice.Add() == 0) {
                    int newDocEntry = int.Parse(Company.GetNewObjectKey());
                    bool isInvoice = (docSubType != 19);
                    bool shouldProceed;
                    if (etapa == 1) {
                        shouldProceed = true;
                    } else {
                        shouldProceed = payInvoice(Company, newDocEntry, isInvoice);
                    }
                    if (shouldProceed) {
                        updateRS.DoQuery("UPDATE " + INVOICES_TABLE + " SET INT_Estado = 'P' WHERE IdFactura = " + idFactura + " AND ExCode = " + exCode + " AND TipoDocumento = " + tipoDoc + " AND Etapa = " + etapa);
                        Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    } else {
                        if (Company.InTransaction) { Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack); }
                        updateRS.DoQuery("UPDATE " + INVOICES_TABLE + " SET INT_Estado = 'E', INT_Error = '" + Company.GetLastErrorDescription().Replace('\'', ' ') + "' WHERE IdFactura = " + idFactura + " AND ExCode = " + exCode + " AND TipoDocumento = " + tipoDoc + " AND Etapa = " + etapa);
                    }
                } else {
                    if (Company.InTransaction) { Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack); }
                    updateRS.DoQuery("UPDATE " + INVOICES_TABLE + " SET INT_Estado = 'E', INT_Error = '" + Company.GetLastErrorDescription().Replace('\'', ' ') + "' WHERE IdFactura = " + idFactura + " AND ExCode = " + exCode + " AND TipoDocumento = " + tipoDoc + " AND Etapa = " + etapa);
                }
            } catch (Exception e) {
                if (Company.InTransaction) { Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack); }
                updateRS.DoQuery("UPDATE " + INVOICES_TABLE + " SET INT_Estado = 'E', INT_Error = '" + e.ToString().Replace('\'', ' ') + "' WHERE IdFactura = " + idFactura + " AND ExCode = " + exCode + " AND TipoDocumento = " + tipoDoc + " AND Etapa = " + etapa);
            }
        }

        private bool payInvoice(SAPbobsCOM.Company Company, int docEntry, bool isInvoice) {
            SAPbobsCOM.Payments payment = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
            SAPbobsCOM.Documents doc = isInvoice ? Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices) : Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes);
            doc.GetByKey(docEntry);

            SAPbobsCOM.Recordset stageOneAccount = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            stageOneAccount.DoQuery("EXEC SEI_ATW_CuentasPago " + doc.UserFields.Fields.Item("U_WebType").Value + ", " + doc.UserFields.Fields.Item("U_ExCode").Value);

            payment.CardCode = doc.CardCode;
            payment.DocDate = doc.DocDate;
            payment.TaxDate = doc.TaxDate;
            payment.DueDate = doc.DocDueDate;
            payment.TransferDate = doc.DocDate;
            payment.DocCurrency = doc.DocCurrency;
            payment.Series = stageOneAccount.Fields.Item("Series").Value;//34 PERU Y CONSULTING, ROOM 32
            payment.TransferAccount = stageOneAccount.Fields.Item("AcctCode").Value;
            payment.Remarks = doc.JournalMemo;
            payment.JournalRemarks = doc.JournalMemo;

            payment.Invoices.InvoiceType = isInvoice ? SAPbobsCOM.BoRcptInvTypes.it_PurchaseInvoice : SAPbobsCOM.BoRcptInvTypes.it_PurchaseCreditNote;
            payment.Invoices.DocEntry = docEntry;
            switch (doc.DocCurrency) {
                case "S/":
                    payment.Invoices.SumApplied = doc.DocTotal;
                    break;
                default:
                    payment.Invoices.AppliedFC = doc.DocTotalFc;
                    break;
            }
            payment.TransferSum = doc.DocCurrency.Equals("S/") ? doc.DocTotal : doc.DocTotalFc;
            return payment.Add() == 0;
        }
    }
}
