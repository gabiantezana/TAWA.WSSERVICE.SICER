using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MigracionTawa.documents {
    class JournalEntries : MDocument {

        private const string JOURNAL_ENTRY_HEADER_SP = "SEI_STW_JournalEntriesHeader";
        private const string JOURNAL_ENTRY_LINES_SP = "SEI_STW_JournalEntriesLines";
        private const string JOURNAL_ENTRY_TABLE = "INT_Tawa_SBO.dbo.CabeceraAsientos";
        private const string JOURNAL_KEY_FIELD = "IdAsiento";

        public JournalEntries(SAPbobsCOM.Company Company)
            : base(Company, JOURNAL_ENTRY_HEADER_SP, JOURNAL_KEY_FIELD) {
        }

        protected override void update(SAPbobsCOM.Company Company, bool successful, string id,string Code) {
            string updateString = "UPDATE " + JOURNAL_ENTRY_TABLE + " SET INT_Estado = '";
            if (successful) {
                updateString += "P' ";
            } else {
                updateString += "E', INT_Error = '" + Company.GetLastErrorDescription().Replace('\'',' ') + "' ";
            }
            updateString += "WHERE IdAsiento = " + id;
            updateRS.DoQuery(updateString);
        }

        protected override bool migrateDocuments(SAPbobsCOM.Company Company, SAPbobsCOM.Recordset migrationRS) {
            SAPbobsCOM.JournalEntries journalEntry = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
            string taxDate = migrationRS.Fields.Item("TaxDate").Value;
            string dueDate = migrationRS.Fields.Item("DueDate").Value;
            string refDate = migrationRS.Fields.Item("ReferenceDate").Value;
            journalEntry.TaxDate = DateTime.ParseExact(taxDate, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
            journalEntry.DueDate = DateTime.ParseExact(dueDate, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
            journalEntry.ReferenceDate = DateTime.ParseExact(refDate, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
            journalEntry.Reference = migrationRS.Fields.Item("Ref1").Value;
            journalEntry.Reference2 = migrationRS.Fields.Item("Ref2").Value;
            journalEntry.Reference3 = migrationRS.Fields.Item("Ref3").Value;
            journalEntry.Memo = migrationRS.Fields.Item("Memo").Value;
            //journalEntry.DocumentType = migrationRS.Fields.Item("Type").Value;

            SAPbobsCOM.Recordset linesRS = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            linesRS.DoQuery(JOURNAL_ENTRY_LINES_SP + " " + migrationRS.Fields.Item("IdAsiento").Value);
            while (!linesRS.EoF) {
                if (!linesRS.BoF) { journalEntry.Lines.Add(); }
                journalEntry.Lines.AccountCode = linesRS.Fields.Item("AccountCode").Value;
                if (String.Equals(linesRS.Fields.Item("FCCurrency").Value, "S/") == false)
                    journalEntry.Lines.FCCurrency = linesRS.Fields.Item("FCCurrency").Value;
                journalEntry.Lines.Debit = linesRS.Fields.Item("Debit").Value;
                journalEntry.Lines.Credit = linesRS.Fields.Item("Credit").Value;
                journalEntry.Lines.FCDebit = linesRS.Fields.Item("FCDebit").Value;
                journalEntry.Lines.FCCredit = linesRS.Fields.Item("FCCredit").Value;
                journalEntry.Lines.Reference1 = linesRS.Fields.Item("Reference1").Value;
                journalEntry.Lines.Reference2 = linesRS.Fields.Item("Reference2").Value;
                journalEntry.Lines.AdditionalReference = linesRS.Fields.Item("Reference3").Value;
                journalEntry.Lines.LineMemo = linesRS.Fields.Item("LineMemo").Value;
                journalEntry.Lines.CostingCode = linesRS.Fields.Item("CostingCode").Value;
                journalEntry.Lines.CostingCode2 = linesRS.Fields.Item("CostingCode2").Value;
                journalEntry.Lines.CostingCode3 = linesRS.Fields.Item("CostingCode3").Value;
                journalEntry.Lines.CostingCode4 = linesRS.Fields.Item("CostingCode4").Value;
                journalEntry.Lines.CostingCode5 = linesRS.Fields.Item("CostingCode5").Value;
                linesRS.MoveNext();
            }
            return journalEntry.Add() == 0;
        }
    }
}
