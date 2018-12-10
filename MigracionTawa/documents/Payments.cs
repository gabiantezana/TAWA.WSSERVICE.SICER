using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MigracionTawa.documents {
    class Payments {

        public void executePayments(SAPbobsCOM.Company Company) {
            SAPbobsCOM.Recordset pendingPayments = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset updateRS = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            pendingPayments.DoQuery("EXEC SEI_STW_AccountPayments");
            while (!pendingPayments.EoF) {
                try {
                    Company.StartTransaction();
                    if (payDocument(Company, pendingPayments)) {
                        updateRS.DoQuery("UPDATE INT_Tawa_SBO.dbo.PagosCuenta SET INT_Estado = 'P' WHERE IdPago = " + pendingPayments.Fields.Item("IdPago").Value);
                        Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    } else {
                        if (Company.InTransaction) { Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack); }
                        updateRS.DoQuery("UPDATE INT_Tawa_SBO.dbo.PagosCuenta SET INT_Estado = 'E', INT_Desc = '" + Company.GetLastErrorDescription().Replace('\'',' ') + "' WHERE IdPago = " + pendingPayments.Fields.Item("IdPago").Value);
                    }
                } catch (Exception) {
                    if (Company.InTransaction) {
                        if (Company.InTransaction) { Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack); }
                        updateRS.DoQuery("UPDATE INT_Tawa_SBO.dbo.PagosCuenta SET INT_Estado = 'E' WHERE IdPago = " + pendingPayments.Fields.Item("IdPago").Value);
                    }
                } finally {
                    pendingPayments.MoveNext();
                }
            }
        }

        private bool payDocument(SAPbobsCOM.Company Company, SAPbobsCOM.Recordset recSetInstance) {
            SAPbobsCOM.Payments incPayment = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
            incPayment.DocType = SAPbobsCOM.BoRcptTypes.rAccount;
            incPayment.DocCurrency = recSetInstance.Fields.Item("Moneda").Value;
            incPayment.ControlAccount = recSetInstance.Fields.Item("CuentaControl").Value;
            incPayment.AccountPayments.AccountCode = recSetInstance.Fields.Item("CuentaDetalle").Value;
            incPayment.AccountPayments.AccountName = recSetInstance.Fields.Item("Nombre").Value;
            incPayment.AccountPayments.SumPaid = recSetInstance.Fields.Item("Monto").Value;
            incPayment.AccountPayments.Decription = recSetInstance.Fields.Item("Memo").Value;
            incPayment.Series = 1; //change

            incPayment.TransferAccount = recSetInstance.Fields.Item("CuentaDetalle").Value;
            incPayment.TransferSum = recSetInstance.Fields.Item("Monto").Value;
            return incPayment.Add() == 0;
        }

    }
}
