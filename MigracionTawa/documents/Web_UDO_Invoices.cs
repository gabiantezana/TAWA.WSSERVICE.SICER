using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MigracionTawa.documents.structs;

namespace MigracionTawa.documents {
    class Web_UDO_Invoices {

        private const string SERVICE_CODE = "-20000";
        private const string INVOICES_HEADER_SP = "SEI_STW_PendingInvoices";
        private const string INVOICES_TABLE = "INT_Tawa_SBO.dbo.FacturasAprobacion";
        private const string INVOICES_KEY = "IdFactura";

        private readonly SAPbobsCOM.GeneralService udoService;
        private readonly Dictionary<ConceptKey, ConceptBody> LcajaChica;
        private readonly Dictionary<ConceptKey, ConceptBody> Lrendiciones;
        private readonly Dictionary<ConceptKey, ConceptBody> Lreembolsos;

        public Web_UDO_Invoices(SAPbobsCOM.Company Company) {
            udoService = Company.GetCompanyService().GetGeneralService(SERVICE_CODE);
            LcajaChica = new Dictionary<ConceptKey, ConceptBody>();
            Lrendiciones = new Dictionary<ConceptKey, ConceptBody>();
            Lreembolsos = new Dictionary<ConceptKey, ConceptBody>();
        }

        public void migrate(SAPbobsCOM.Company Company) {
            SAPbobsCOM.Recordset migrationRS = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            migrationRS.DoQuery("[SEI_STW_WebUDOInvoices]");
            while (!migrationRS.EoF) {
                int docT = migrationRS.Fields.Item("TipoDocumento").Value;
                int exC = migrationRS.Fields.Item("ExCode").Value;
                int stage = migrationRS.Fields.Item("Etapa").Value;
                double monBrut = migrationRS.Fields.Item("MonBruto").Value;
                double docRate = migrationRS.Fields.Item("DocRate").Value;
                string currency = migrationRS.Fields.Item("DocCurrency").Value;
                ConceptKey key = new ConceptKey(exC, stage, migrationRS.Fields.Item("CostingCode").Value, migrationRS.Fields.Item("CostingCode2").Value, migrationRS.Fields.Item("CostingCode3").Value, migrationRS.Fields.Item("CostingCode4").Value, migrationRS.Fields.Item("CostingCode5").Value, migrationRS.Fields.Item("DocCurrency").Value, migrationRS.Fields.Item("JournalMemo").Value);
                switch (docT) {
                    case 3:
                        if (LcajaChica.ContainsKey(key)) {
                            if (!currency.Equals(key.currency)) { monBrut = currency.Equals("S/") ? monBrut / docRate : monBrut * docRate; }
                            LcajaChica[key].incSum(monBrut);
                        } else {
                            ConceptBody newBody = new ConceptBody(migrationRS.Fields.Item("DocDate").Value, migrationRS.Fields.Item("MonBruto").Value);
                            LcajaChica.Add(key, newBody);
                        }
                        break;
                    case 4:
                        if (Lrendiciones.ContainsKey(key)) {
                            if (!currency.Equals(key.currency)) { monBrut = currency.Equals("S/") ? monBrut / docRate : monBrut * docRate; }
                            Lrendiciones[key].incSum(monBrut);
                        } else {
                            ConceptBody newBody = new ConceptBody(migrationRS.Fields.Item("DocDate").Value, migrationRS.Fields.Item("MonBruto").Value);
                            Lrendiciones.Add(key, newBody);
                        }
                        break;
                    case 5:
                        if (Lreembolsos.ContainsKey(key)) {
                            if (!currency.Equals(key.currency)) { monBrut = currency.Equals("S/") ? monBrut / docRate : monBrut * docRate; }
                            Lreembolsos[key].incSum(monBrut);
                        } else {
                            ConceptBody newBody = new ConceptBody(migrationRS.Fields.Item("DocDate").Value, migrationRS.Fields.Item("MonBruto").Value);
                            Lreembolsos.Add(key, newBody);
                        }
                        break;
                }
                migrationRS.MoveNext();
            }

            foreach (KeyValuePair<ConceptKey, ConceptBody> entry in LcajaChica) {
                singleMigration(Company, entry.Key, entry.Value, 3);
            }

            foreach (KeyValuePair<ConceptKey, ConceptBody> entry in Lrendiciones) {
                entregasARendir(Company, entry.Key, entry.Value);
            }

            foreach (KeyValuePair<ConceptKey, ConceptBody> entry in Lreembolsos) {
                singleMigration(Company, entry.Key, entry.Value, 5);
            }
        }

        private void singleMigration(SAPbobsCOM.Company Company, ConceptKey key, ConceptBody body, int docType) {
            Company.StartTransaction();
            SAPbobsCOM.Recordset generalInfoRS = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try {
                SAPbobsCOM.GeneralData singleInvoice = udoService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                DateTime docDate = body.getDate();
                generalInfoRS.DoQuery("EXEC [SEI_STW_AccountInfo] '" + key.currency + "', '" + body.parsedDate + "', '" + key.costingCode5 + "'");
                //singleInvoice.SetProperty("U_CardCode", key.exCode.ToString());
                singleInvoice.SetProperty("U_InvDocEntry", key.exCode.ToString());
                singleInvoice.SetProperty("U_LineMemo", key.asunto);
                singleInvoice.SetProperty("U_DocCurrency", key.currency);
                singleInvoice.SetProperty("U_CostingCode", key.costingCode);
                singleInvoice.SetProperty("U_CostingCode2", key.costingCode2);
                singleInvoice.SetProperty("U_CostingCode3", key.costingCode3);
                singleInvoice.SetProperty("U_CostingCode4", key.costingCode4);
                singleInvoice.SetProperty("U_CostingCode5", key.costingCode5);
                singleInvoice.SetProperty("U_Year", docDate.Year);
                singleInvoice.SetProperty("U_Month", docDate.Month);
                singleInvoice.SetProperty("U_MonBruto", body.getTotalSum());
                singleInvoice.SetProperty("U_DocType", docType);
                singleInvoice.SetProperty("U_DocDate", docDate);
                singleInvoice.SetProperty("U_TipoCambio", generalInfoRS.Fields.Item("Rate").Value);
                singleInvoice.SetProperty("U_AccountCode", generalInfoRS.Fields.Item("U_Asociada").Value);
                singleInvoice.SetProperty("U_DetAcctCode", generalInfoRS.Fields.Item("U_Detalle").Value);
                singleInvoice.SetProperty("U_FactRefact", generalInfoRS.Fields.Item("Multiplier").Value);
                singleInvoice.SetProperty("U_Usable", "Y");
                singleInvoice.SetProperty("U_Tipo", key.costingCode3);
                generalInfoRS.DoQuery("SELECT CardCode FROM INT_Tawa_SBO.dbo.FacturasWeb WHERE Etapa = 1 AND ExCode = '" + key.exCode + "' AND TipoDocumento = '" + docType + "' AND JournalMemo = '" + key.asunto + "'");
                singleInvoice.SetProperty("U_CardCode", generalInfoRS.Fields.Item("CardCode").Value);
                udoService.Add(singleInvoice);
                bool pending = ((docType == 4) && (key.etapa == 1));
                generalInfoRS.DoQuery("UPDATE INT_Tawa_SBO.dbo.FacturasWeb SET INT_Ref1 = '" + (pending ? "A" : "P") + "' WHERE ExCode = " + key.exCode + " AND Etapa = " + key.etapa + " AND TipoDocumento = " + docType + " AND JournalMemo = '" + key.asunto + "'");
                Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
            } catch (Exception) {
                if (Company.InTransaction) { Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack); }
                generalInfoRS.DoQuery("UPDATE INT_Tawa_SBO.dbo.FacturasWeb SET INT_Ref1 = 'E' WHERE ExCode = " + key.exCode + " AND Etapa = " + key.etapa + " AND TipoDocumento = " + docType + " AND JournalMemo = '" + key.asunto + "'");
            }
        }

        private void entregasARendir(SAPbobsCOM.Company Company, ConceptKey key, ConceptBody body) {
            SAPbobsCOM.Recordset generalPurposeRS = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try {
                if (key.etapa == 1) {
                    generalPurposeRS.DoQuery("EXEC [SEI_STW_PrimeraEtapa] " + key.exCode);
                    int count = generalPurposeRS.Fields.Item("Count").Value;
                    if (count == 0) {
                        singleMigration(Company, key, body, 4);
                    } else {
                        Company.StartTransaction();
                        generalPurposeRS.DoQuery("EXEC [SEI_STW_RendicionesSegundaEtapa] " + key.exCode + ", '" + key.currency + "'");
                        if (generalPurposeRS.RecordCount == 0) { 
                            Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack); 
                            return; 
                        }
                        double dTotal = generalPurposeRS.Fields.Item("DocTotal").Value / generalPurposeRS.Fields.Item("Rate").Value;
                        if (dTotal == body.getTotalSum()) {
                            generalPurposeRS.DoQuery("UPDATE INT_Tawa_SBO.dbo.FacturasWeb SET INT_Ref1 = 'P' WHERE ExCode = " + key.exCode + " AND Etapa = " + key.etapa + " AND TipoDocumento = 4 AND JournalMemo = '" + key.asunto + "'");
                            Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                        }else{
                            //bool exceedsSum = dTotal > body.getTotalSum();
                            //double newDocumentSum = Math.Abs(dTotal - body.getTotalSum());
                            //if (rectificationDocument(Company,ref generalPurposeRS, body, key, exceedsSum, newDocumentSum)) {
                                generalPurposeRS.DoQuery("UPDATE INT_Tawa_SBO.dbo.FacturasWeb SET INT_Ref1 = 'P' WHERE ExCode = " + key.exCode + " AND Etapa = " + key.etapa + " AND TipoDocumento = 4 AND JournalMemo = '" + key.asunto + "'");
                                Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                            //} else {
                            //    if (Company.InTransaction) { Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack); }
                            //    generalPurposeRS.DoQuery("UPDATE INT_Tawa_SBO.dbo.FacturasWeb SET INT_Ref1 = 'E', INT_Error = '" + Company.GetLastErrorDescription().Replace('\'', ' ') + "' WHERE ExCode = " + key.exCode + " AND Etapa = " + key.etapa + " AND TipoDocumento = 4");
                            //}
                        }
                    }
                } else {
                    singleMigration(Company, key, body, 4);
                }
            } catch (Exception) {
                if (Company.InTransaction) { Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack); }
                generalPurposeRS.DoQuery("UPDATE INT_Tawa_SBO.dbo.FacturasWeb SET INT_Ref1 = 'E' WHERE ExCode = " + key.exCode + " AND Etapa = " + key.etapa + " AND TipoDocumento = 4 AND JournalMemo = '" + key.asunto + "'");
            }
        }

        private bool rectificationDocument(SAPbobsCOM.Company Company,ref SAPbobsCOM.Recordset generalPurposeRS, ConceptBody body, ConceptKey key, bool exceedsSum,double sum) {
            SAPbobsCOM.Documents newInvoice = exceedsSum ? Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices) : Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes);
            newInvoice.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service;
            newInvoice.DocCurrency = key.currency;
            newInvoice.DocDate = body.getDate();
            newInvoice.DocDueDate = body.getDate();
            newInvoice.TaxDate = body.getDate();
            newInvoice.Comments = key.asunto;
            newInvoice.FolioPrefixString = "DER";
            newInvoice.FolioNumber = generalPurposeRS.Fields.Item("Folio").Value;
            newInvoice.Series = generalPurposeRS.Fields.Item("Series").Value;
            newInvoice.UserFields.Fields.Item("U_MSS_REFACT_COT").Value = "SI";
            newInvoice.UserFields.Fields.Item("U_MSS_Refacturado").Value = "N";
            newInvoice.UserFields.Fields.Item("U_ExCode").Value = key.exCode;
            newInvoice.UserFields.Fields.Item("U_WebType").Value = 4;
            newInvoice.UserFields.Fields.Item("U_Etapa").Value = key.etapa;
            newInvoice.UserFields.Fields.Item("U_BPP_MDTD").Value = "01";

            newInvoice.Lines.TaxCode = "IGV_EXE";
            newInvoice.Lines.LineTotal = sum;
            newInvoice.Lines.CostingCode = key.costingCode;
            newInvoice.Lines.CostingCode2 = key.costingCode2;
            newInvoice.Lines.CostingCode3 = key.costingCode3;
            newInvoice.Lines.CostingCode4 = key.costingCode4;
            newInvoice.Lines.CostingCode5 = key.costingCode5;

            generalPurposeRS.DoQuery("EXEC [SEI_STW_DatosEREtapa1] " + key.exCode);
            newInvoice.CardCode = generalPurposeRS.Fields.Item("CardCode").Value;
            newInvoice.JournalMemo = generalPurposeRS.Fields.Item("JournalMemo").Value;
            newInvoice.ControlAccount = generalPurposeRS.Fields.Item("U_Asociada").Value;
            newInvoice.Lines.AccountCode = generalPurposeRS.Fields.Item("U_Detalle").Value;
            newInvoice.Lines.ItemDescription = generalPurposeRS.Fields.Item("JournalMemo").Value;
            return newInvoice.Add() == 0;
        }
    }
}
