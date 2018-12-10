using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MigracionTawa.documents {
    abstract class MDocument : IDisposable {
        protected readonly string migrationSP;
        protected readonly string keyField;
        protected SAPbobsCOM.Recordset updateRS;

        protected MDocument(SAPbobsCOM.Company Company, string migrationStoredProcedure, string KeyField) {
            updateRS = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            migrationSP = migrationStoredProcedure;
            keyField = KeyField;
        }

        public void migrateBP(SAPbobsCOM.Company Company) {
            SAPbobsCOM.Recordset migrationRS = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            migrationRS.DoQuery(migrationSP);
            while (!migrationRS.EoF) {
                Company.StartTransaction();
                string currentDocEntry = migrationRS.Fields.Item(keyField).Value;
                string code = migrationRS.Fields.Item("Code").Value;
                try {
                    if (migrateDocuments(Company, migrationRS)) {
                        update(Company, true, currentDocEntry.ToString(), code);
                        Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    } else {
                        Company.GetLastErrorDescription();
                        if (Company.InTransaction) { Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack); }
                        update(Company, false, currentDocEntry.ToString(), code);
                    }

                } catch (Exception) {
                    if (Company.InTransaction) { Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack); }
                    update(Company, false, currentDocEntry, code);
                }
                migrationRS.MoveNext();
            }
            if (Company.InTransaction) { Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack); }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(migrationRS);
            migrationRS = null;
        }


        public void migrate(SAPbobsCOM.Company Company) {
            SAPbobsCOM.Recordset migrationRS = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            migrationRS.DoQuery(migrationSP);
            while (!migrationRS.EoF) {
                Company.StartTransaction();
                int currentDocEntry = migrationRS.Fields.Item(keyField).Value;
                string Code = migrationRS.Fields.Item("Code").Value;
                try {
                    if (migrateDocuments(Company, migrationRS)) {
                        update(Company, true, currentDocEntry.ToString(), Code);
                        Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    } else {
                        Company.GetLastErrorDescription();
                        if (Company.InTransaction) { Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack); }
                        update(Company, false, currentDocEntry.ToString(), Code);
                    }

                } catch (Exception) {
                    if (Company.InTransaction) { Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack); }
                    update(Company, false, currentDocEntry.ToString(), Code);
                }
                migrationRS.MoveNext();
            }
            if (Company.InTransaction) { Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack); }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(migrationRS);
            migrationRS = null;
        }

        protected abstract void update(SAPbobsCOM.Company Company, bool successful, string id, string Code);
        protected abstract bool migrateDocuments(SAPbobsCOM.Company Company, SAPbobsCOM.Recordset migrationRS);

        public void Dispose() {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(updateRS);
            updateRS = null;
        }
    }
}
