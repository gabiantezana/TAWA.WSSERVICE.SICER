using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MigracionTawa.connection {
    class Connection {
        private readonly Dictionary<string, SAPbobsCOM.Company> companies;
        private const string CONNECTION_FILE = "\\conexion.xml";
        private const string FILE_OUTPUT_DIRECTION = "C:\\";

        public Connection() {
            companies = new Dictionary<string, SAPbobsCOM.Company>();
        }

        public void initializeConnections() {
            System.Xml.Linq.XDocument connectionXML = System.Xml.Linq.XDocument.Load(AppDomain.CurrentDomain.BaseDirectory + CONNECTION_FILE);
            var xmlNodes = from header in connectionXML.Descendants("Company")
                           select new {
                               CompanyCode = header.Element("DBCompany").Value,
                               Server = header.Element("Server").Value,
                               DBUser = header.Element("DBUser").Value,
                               DBPassword = header.Element("DBPassword").Value,
                               SBOUser = header.Element("SBOUser").Value,
                               SBOPassword = header.Element("SBOPassword").Value
                           };
            foreach (var xmlNode in xmlNodes) {
               SAPbobsCOM.Company  company = new SAPbobsCOM.Company();
                company.CompanyDB = xmlNode.CompanyCode;
                company.Server = xmlNode.Server;
                company.DbUserName = xmlNode.DBUser;
                company.DbPassword = xmlNode.DBPassword;
                company.UserName = xmlNode.SBOUser;
                company.Password = xmlNode.SBOPassword;
                company.language = SAPbobsCOM.BoSuppLangs.ln_Spanish;
                company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;
                if (company.Connect() == 0) {
                    companies.Add(company.CompanyDB, company);
                } else {
                    company.GetLastErrorDescription();
                }
            }
        }

        public SAPbobsCOM.Company getCompany(string companyCode) {
            return companies[companyCode];
        }

        public void Dispose() {
            foreach (KeyValuePair<String, SAPbobsCOM.Company> company in companies) {
                if (company.Value.Connected) company.Value.Disconnect();
            }
            companies.Clear();
        }

        private void notify(string message, string type) {
            try {
                string currentDate = DateTime.Now.ToString("yyyy/MM/dd_HH:mm");
                System.IO.File.WriteAllLines(FILE_OUTPUT_DIRECTION + type + currentDate + ".txt", new String[] { message });
            } catch (Exception e) {
                notify(FILE_OUTPUT_DIRECTION + "UnhandledException.txt", e.ToString());
            }
        }

        public HashSet<string> companiesConnected() {
            HashSet<string> retVal = new HashSet<string>();
            foreach (KeyValuePair<string, SAPbobsCOM.Company> company in companies) {
                retVal.Add(company.Key);
            }
            return retVal;
        }

        public int getCount() { return companies.Count; }

    }
}
