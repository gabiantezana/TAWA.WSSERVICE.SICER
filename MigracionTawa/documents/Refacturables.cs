using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MigracionTawa.documents {
    class Refacturables {

        private const string SERVICE_CODE = "-20000";
        private const string SP_REFACTURABLES = "EXEC SEI_STW_Refacturables";
        private const int PURCHASE_INVOICES = 18;
        private const int PURCHASE_CREDIT_NOTES = 19;

        private readonly SAPbobsCOM.GeneralService udoService;

        public Refacturables(SAPbobsCOM.Company Company) {
            udoService = Company.GetCompanyService().GetGeneralService(SERVICE_CODE);
        }

        public void ingresarRefacturables(SAPbobsCOM.Company Company) {
            SAPbobsCOM.Recordset refRS = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            refRS.DoQuery(SP_REFACTURABLES);
            while (!refRS.EoF) {
                Company.StartTransaction();
                if (migrate(Company, refRS)) {
                    Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                } else {
                    Company.GetLastErrorDescription();
                    if (Company.InTransaction) { Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack); }
                }
                refRS.MoveNext();                
            }
        }

        private bool migrate(SAPbobsCOM.Company Company, SAPbobsCOM.Recordset recSet) {
            try {
                int objType = recSet.Fields.Item("ObjType").Value;
                int docEntry = recSet.Fields.Item("DocEntry").Value;
                SAPbobsCOM.Documents refDocument;
                switch (objType) {
                    case PURCHASE_INVOICES:
                        refDocument = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                        break;
                    case PURCHASE_CREDIT_NOTES:
                        refDocument = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes);
                        break;
                    default:
                        return false;
                }
                if (refDocument.GetByKey(docEntry)) {
                    SAPbobsCOM.Recordset sRate = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    sRate.DoQuery("SELECT Rate FROM ORTT WHERE Currency = 'US$' AND RateDate = '" + refDocument.DocDate.ToString("yyyyMMdd") + "'");
                    double monBruto = refDocument.DocCurrency.Equals("S/") ? refDocument.DocTotal - refDocument.VatSum : refDocument.DocTotalFc - refDocument.VatSumFc;
                    SAPbobsCOM.GeneralData singleRecord = udoService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                    singleRecord.SetProperty("U_CardCode", refDocument.CardCode);
                    singleRecord.SetProperty("U_InvDocEntry", refDocument.DocEntry);
                    singleRecord.SetProperty("U_LineMemo", refDocument.Comments);
                    singleRecord.SetProperty("U_DocCurrency", refDocument.DocCurrency);
                    singleRecord.SetProperty("U_CostingCode", refDocument.Lines.CostingCode);
                    singleRecord.SetProperty("U_CostingCode2", refDocument.Lines.CostingCode2);
                    singleRecord.SetProperty("U_CostingCode3", refDocument.Lines.CostingCode3);
                    singleRecord.SetProperty("U_CostingCode4", refDocument.Lines.CostingCode4);
                    singleRecord.SetProperty("U_CostingCode5", refDocument.Lines.CostingCode5);
                    singleRecord.SetProperty("U_Year", refDocument.DocDate.Year);
                    singleRecord.SetProperty("U_Month", refDocument.DocDate.Month);
                    singleRecord.SetProperty("U_MonBruto", objType == 18 ? monBruto : -monBruto);
                    singleRecord.SetProperty("U_DocType", 2);
                    singleRecord.SetProperty("U_DocDate", refDocument.DocDate);
                    singleRecord.SetProperty("U_TipoCambio", sRate.Fields.Item("Rate").Value);
                    singleRecord.SetProperty("U_AccountCode", refDocument.ControlAccount);
                    singleRecord.SetProperty("U_DetAcctCode", refDocument.Lines.AccountCode);
                    singleRecord.SetProperty("U_FactRefact", recSet.Fields.Item("Multiplier").Value);
                    singleRecord.SetProperty("U_Usable", "Y");
                    singleRecord.SetProperty("U_Tipo", refDocument.Lines.CostingCode3);
                    udoService.Add(singleRecord);
                    refDocument.UserFields.Fields.Item("U_MSS_Refacturado").Value = "Y";
                    return refDocument.Update() == 0;
                } else {
                    Company.GetLastErrorDescription();
                    return false;
                }
            } catch (Exception) {
                return false;
            }
        }
    }
}
