using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Timers;
using MigracionTawa.connection;
using MigracionTawa.documents;
using MigracionTawa.documents.structs;

namespace MigracionTawa {
    public partial class MainTasks : ServiceBase {

        private const double INITIAL_TIME = 5000.0d;
        private const double CYCLE_INTERVAL = 5000.0d;
        private Timer trigger;

        public MainTasks() {
            InitializeComponent();
        }

        protected override void OnStart(string[] args) {
            trigger = new Timer(INITIAL_TIME);
            trigger.Interval = CYCLE_INTERVAL;
            trigger.AutoReset = false;
            trigger.Elapsed += new System.Timers.ElapsedEventHandler(Tasks);
            trigger.Enabled = true;
        }

        protected override void OnStop() {
            trigger.Dispose();
        }

        private void Tasks(object sender, ElapsedEventArgs e) {
            trigger.Enabled = false;
            Connection con = new Connection();
            try {
                con.initializeConnections();
                HashSet<string> companies = con.companiesConnected();
                foreach (string company in companies) {
                    SAPbobsCOM.Company Company = con.getCompany(company);

                    ////BusinessPartners bp = new BusinessPartners(Company);
                    ////bp.migrateBP(Company);

                    ////JournalEntries je = new JournalEntries(Company);
                    ////je.migrate(Company);

                    ////Invoices rInv = new Invoices(Company);
                    ////rInv.migrate(Company);

                    ////WebInvoices webInv = new WebInvoices();
                    ////webInv.migrate(Company);

                    ////SalesInvoices salesInv = new SalesInvoices(Company);
                    ////salesInv.migrate(Company);

                    CorreoLideres mailLideres = new CorreoLideres();
                    mailLideres.EnviarCorreoLideres(Company);
                    mailLideres.EnviarCorreoPagosEfectuados(Company);

                    //Refacturables AddOn

                    ////Web_UDO_Invoices udoInv = new Web_UDO_Invoices(Company);
                    ////udoInv.migrate(Company);

                    ////Refacturables refact = new Refacturables(Company);
                    ////refact.ingresarRefacturables(Company);

                    ////UDO_Invoices inv = new UDO_Invoices(Company);
                    ////inv.migrate(Company);

                    //Payments payments = new Payments();
                    //payments.executePayments(Company);

                    //delUDO(Company);
                }
                con.Dispose();
            } catch (Exception) {
                con.Dispose();
            } finally {
                GC.WaitForPendingFinalizers();
                GC.Collect();
                trigger.Enabled = true;
            }
        }

        private void delUDO(SAPbobsCOM.Company Company) {
            SAPbobsCOM.Recordset gRS = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            gRS.DoQuery("SELECT TOP 100 DocEntry FROM [@SEI_UREF]");
            SAPbobsCOM.GeneralService invServ = Company.GetCompanyService().GetGeneralService("-20000");
            while (!gRS.EoF) {
                try {
                    SAPbobsCOM.GeneralDataParams sP = invServ.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                    sP.SetProperty("DocEntry", gRS.Fields.Item("DocEntry").Value);
                    invServ.Delete(sP);
                } catch (Exception) { }
                gRS.MoveNext();
            }
        }

    }
}
