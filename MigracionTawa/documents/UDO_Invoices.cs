using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MigracionTawa.documents {
    class UDO_Invoices : MDocument, IDisposable {

        private const string SERVICE_CODE = "-20000";
        private const string INVOICES_HEADER_SP = "SEI_STW_PendingInvoices";
        private const string INVOICES_TABLE = "INT_Tawa_SBO.dbo.FacturasAprobacion";
        private const string INVOICES_KEY = "IdFactura";

        private readonly SAPbobsCOM.GeneralService invoiceService;

        public UDO_Invoices(SAPbobsCOM.Company Company)
            : base(Company, INVOICES_HEADER_SP, INVOICES_KEY) {
            invoiceService = Company.GetCompanyService().GetGeneralService(SERVICE_CODE);
        }

        protected override bool migrateDocuments(SAPbobsCOM.Company Company, SAPbobsCOM.Recordset migrationRS) {
            SAPbobsCOM.GeneralData invoice = invoiceService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
            string invDate = migrationRS.Fields.Item("INT_Fecha").Value;
            int concType = migrationRS.Fields.Item("Concepto").Value;
            DateTime pDate = DateTime.ParseExact(invDate, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
            invoice.SetProperty("U_MSS_GRPFACT", migrationRS.Fields.Item("CO_GRUP_FACT").Value);
            invoice.SetProperty("U_CardCode", migrationRS.Fields.Item("CardCode").Value);
            invoice.SetProperty("U_Year", migrationRS.Fields.Item("Año").Value);
            invoice.SetProperty("U_Month", migrationRS.Fields.Item("Mes").Value);
            invoice.SetProperty("U_TipoCambio", migrationRS.Fields.Item("TipoCambio").Value);
            invoice.SetProperty("U_Aprobador", migrationRS.Fields.Item("Aprobador").Value);
            invoice.SetProperty("U_Tipo", migrationRS.Fields.Item("CostingCode3").Value);
            invoice.SetProperty("U_RemuAfecta", migrationRS.Fields.Item("IM_REMU_AFEC").Value);
            invoice.SetProperty("U_RemuInaf", migrationRS.Fields.Item("IM_REMU_INAF").Value);
            invoice.SetProperty("U_MonBruto", migrationRS.Fields.Item("IM_BRUT_FACT").Value);
            invoice.SetProperty("U_MonNeto", migrationRS.Fields.Item("IM_NETO_FACT").Value);
            invoice.SetProperty("U_MonAnt", migrationRS.Fields.Item("IM_BRUT_ANTI").Value);
            invoice.SetProperty("U_MonIGV", migrationRS.Fields.Item("IM_BRUT_IGV").Value);
            invoice.SetProperty("U_ImpRefact", migrationRS.Fields.Item("SAP_IMPREFAC").Value);
            invoice.SetProperty("U_AccountCode", migrationRS.Fields.Item("AccountCode").Value); //lel
            invoice.SetProperty("U_DocType", migrationRS.Fields.Item("Concepto").Value); //top lel
            invoice.SetProperty("U_CostingCode", migrationRS.Fields.Item("CostingCode").Value);
            invoice.SetProperty("U_CostingCode2", migrationRS.Fields.Item("CostingCode2").Value);
            invoice.SetProperty("U_CostingCode3", migrationRS.Fields.Item("CostingCode3").Value);
            invoice.SetProperty("U_CostingCode4", migrationRS.Fields.Item("CostingCode4").Value);
            invoice.SetProperty("U_CostingCode5", migrationRS.Fields.Item("CostingCode5").Value);
            invoice.SetProperty("U_DocCurrency", migrationRS.Fields.Item("DocCurrency").Value);
            invoice.SetProperty("U_DetAcctCode",migrationRS.Fields.Item("AccountCodeDet").Value);
            invoice.SetProperty("U_FactRefact", migrationRS.Fields.Item("FactorRefacturacion").Value); //Crear
            invoice.SetProperty("U_STPLAN_OTRO", migrationRS.Fields.Item("ST_PLAN_OTRO").Value);
            invoice.SetProperty("U_COPLAN_OTRO", migrationRS.Fields.Item("CO_PLAN_OTRO").Value);
            invoice.SetProperty("U_FAOTRO_AFEC", migrationRS.Fields.Item("FA_OTRO_AFEC").Value);
            invoice.SetProperty("U_FAOTRO_INAF", migrationRS.Fields.Item("FA_OTRO_INAF").Value);
            invoice.SetProperty("U_FAOTRO_VACA", migrationRS.Fields.Item("FA_OTRO_VACA").Value);
            invoice.SetProperty("U_LineMemo", migrationRS.Fields.Item("NO_GLOS_SAP1").Value);
            invoice.SetProperty("U_DocDate", pDate);
            invoice.SetProperty("U_Usable", "Y");
            invoice.SetProperty("U_InvDocEntry", migrationRS.Fields.Item("DocEntry").Value);
            invoiceService.Add(invoice);
            return true;
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
    }
}
