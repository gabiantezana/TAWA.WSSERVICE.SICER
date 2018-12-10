using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MigracionTawa.documents {

    class CorreoLideres {

        private const string CorreoPendienteLideres = "CorreoPendienteEnvio";
        private const string CuerpoMensajeLideres = "CuerpoCorreoLideres";
        private const string CorreoDestinatarioLideres = "DestinatarioCorreo";
        private const string CorreoPendienteLideresUpdate = "CorreoPendienteUpdate";

        private const string CorreoPendientePagosEfectuados = "CorreosPagosEfectuadosPendienteEnvio";
        private const string CuerpoMensajePagosEfectuados = "CuerpoCorreoPagosEfectuados ";
        private const string CorreoDestinatarioPagosEfectuados = "DestinatarioCorreo ";
        private const string CorreoPendientePagosEfectuadosUpdate = "CorreoPendientePagosEfectuadosUpdate";
       


        public void EnviarCorreoLideres(SAPbobsCOM.Company Company) {

            try {

                SAPbobsCOM.Recordset rsObtenerCorreoPendiente = null;
                SAPbobsCOM.Recordset rsObtenerCorreoCuerpo = null;
                SAPbobsCOM.Recordset rsObtenerCorreoDestinatario = null;
                SAPbobsCOM.Recordset rsObtenerCorreoPendienteUpdate = null;


                rsObtenerCorreoPendiente = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rsObtenerCorreoPendiente.DoQuery(CorreoPendienteLideres);



                if (rsObtenerCorreoPendiente.RecordCount > 0) {

                    String BaseDatos = rsObtenerCorreoPendiente.Fields.Item("BaseDatos").Value;
                    String docEntry = rsObtenerCorreoPendiente.Fields.Item("DocEntry").Value;
                    String ObjType = rsObtenerCorreoPendiente.Fields.Item("ObjType").Value;
                    String NombreSolicitante = rsObtenerCorreoPendiente.Fields.Item("NombreSolicitante").Value;
                    String DocnNum = rsObtenerCorreoPendiente.Fields.Item("DocNum").Value;
                    String Status = rsObtenerCorreoPendiente.Fields.Item("Status").Value;
                    String FechaCreacion = Convert.ToString(rsObtenerCorreoPendiente.Fields.Item("FechaCreacion").Value);
                    String Area = rsObtenerCorreoPendiente.Fields.Item("Area").Value;
                    String Email = rsObtenerCorreoPendiente.Fields.Item("Email").Value;

                    rsObtenerCorreoCuerpo = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    rsObtenerCorreoCuerpo.DoQuery(CuerpoMensajeLideres);

                    String Cabecera = rsObtenerCorreoCuerpo.Fields.Item("Cabecera").Value;
                    String Cuerpo = rsObtenerCorreoCuerpo.Fields.Item("Cuerpo").Value;
                    String Solicitante = rsObtenerCorreoCuerpo.Fields.Item("Solicitante").Value;
                    String DocNum = rsObtenerCorreoCuerpo.Fields.Item("DocNum").Value;
                    String Fecha = rsObtenerCorreoCuerpo.Fields.Item("Fecha").Value;
                    String AreaCuerpo = rsObtenerCorreoCuerpo.Fields.Item("Area").Value;
                    String Pie = rsObtenerCorreoCuerpo.Fields.Item("Pie").Value;

                    String Mensaje = "";
                    String Destino = "";
                    String NombreLider = "";

                    rsObtenerCorreoDestinatario = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    rsObtenerCorreoDestinatario.DoQuery(CorreoDestinatarioLideres + " '" + NombreSolicitante + "'");
                    Destino = rsObtenerCorreoDestinatario.Fields.Item("U_MSS_Correo").Value;
                    NombreLider = rsObtenerCorreoDestinatario.Fields.Item("U_MSS_Lider").Value;

                    Mensaje = Cabecera + NombreLider + "\n" + "\n" +
                        Cuerpo + "\n" + "\n" +
                        Solicitante + NombreSolicitante + "\n" +
                        DocNum + DocnNum + "\n" +
                        Fecha + FechaCreacion + "\n" +
                        //AreaCuerpo + Area + "\n" +
                        Pie + "\n";
                        //NombreSolicitante;

                   

                   
                        if (MensajeMail(Mensaje, "Creación de Solicitud de Compra N° " + DocnNum, Destino)) {
                            rsObtenerCorreoPendienteUpdate = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            rsObtenerCorreoPendienteUpdate.DoQuery(CorreoPendienteLideresUpdate + " '" + docEntry + "'");
                        }

                }

            } catch (Exception) {

                throw;
            }

        }

        public void EnviarCorreoPagosEfectuados(SAPbobsCOM.Company Company) {

            try {

                SAPbobsCOM.Recordset rsObtenerCorreoPendientePagosEfectuados = null;
                SAPbobsCOM.Recordset rsObtenerCorreoCuerpoPagosEfectuados = null;
                SAPbobsCOM.Recordset rsObtenerCorreoDestinatarioPagosEfectuados = null;
                SAPbobsCOM.Recordset rsObtenerCorreoPendientePagosEfectuadosUpdate = null;


                rsObtenerCorreoPendientePagosEfectuados = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rsObtenerCorreoPendientePagosEfectuados.DoQuery(CorreoPendientePagosEfectuados);



                if (rsObtenerCorreoPendientePagosEfectuados.RecordCount > 0) {

                    String BaseDatos = rsObtenerCorreoPendientePagosEfectuados.Fields.Item("BaseDatos").Value;
                    String docEntry = rsObtenerCorreoPendientePagosEfectuados.Fields.Item("DocEntry").Value;
                    String CardCode = rsObtenerCorreoPendientePagosEfectuados.Fields.Item("CardCode").Value;
                    String CardName = rsObtenerCorreoPendientePagosEfectuados.Fields.Item("CardName").Value;
                    String Total = rsObtenerCorreoPendientePagosEfectuados.Fields.Item("Total").Value;
                    String Estado = rsObtenerCorreoPendientePagosEfectuados.Fields.Item("Estado").Value;
                    String Email = rsObtenerCorreoPendientePagosEfectuados.Fields.Item("Email").Value;

                    rsObtenerCorreoCuerpoPagosEfectuados = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    rsObtenerCorreoCuerpoPagosEfectuados.DoQuery(CuerpoMensajePagosEfectuados);

                    String Cabecera = rsObtenerCorreoCuerpoPagosEfectuados.Fields.Item("Cabecera").Value;
                    String Cuerpo1 = rsObtenerCorreoCuerpoPagosEfectuados.Fields.Item("Cuerpo1").Value;
                    String Cuerpo2 = rsObtenerCorreoCuerpoPagosEfectuados.Fields.Item("Cuerpo2").Value;
                    String Cuerpo3 = rsObtenerCorreoCuerpoPagosEfectuados.Fields.Item("Cuerpo3").Value;
                    String Cuerpo4 = rsObtenerCorreoCuerpoPagosEfectuados.Fields.Item("Cuerpo4").Value;
                    String Soles = rsObtenerCorreoCuerpoPagosEfectuados.Fields.Item("Soles").Value;
                    String Dolares = rsObtenerCorreoCuerpoPagosEfectuados.Fields.Item("Dolares").Value;

                    String Mensaje = "";
                    String Destino = Email ;
                    

                        Mensaje = Cabecera + CardName + "\n" + "\n" +
                        Cuerpo1 + Total + "\n" +
                        Cuerpo2 + "\n" +
                        Cuerpo3 + "\n" +
                        Cuerpo4 ;
                     
                   


                        if (MensajeMail(Mensaje, "Pago efectuado de la empresa TAWA" + docEntry , Destino)) {
                            rsObtenerCorreoPendientePagosEfectuadosUpdate = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            rsObtenerCorreoPendientePagosEfectuadosUpdate.DoQuery(CorreoPendientePagosEfectuadosUpdate  + " '" + docEntry + "'");
                        }

                }

            } catch (Exception ex) {

                throw;
            }

        }

        private bool MensajeMail(string Cuerpo, string Asunto, String Destino) {
            if (Destino.Trim() != "") {
                System.Net.Mail.MailMessage correo = new System.Net.Mail.MailMessage();
                String email_body = "";
                correo.From = new System.Net.Mail.MailAddress("compras.peru@tawa.com.pe");
                correo.To.Add("compras@Grupotawa.com");
                correo.To.Add("arturojunior.rodriguez@gmail.com");
                correo.To.Add(Destino.Trim());
                correo.Subject = Asunto;
                email_body = Cuerpo + "";
                correo.Body = email_body;
                correo.Priority = System.Net.Mail.MailPriority.Normal;
                System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient();
                smtp.Host = "mailhost1.tawa.com.pe";
                smtp.EnableSsl = false;

                try {
                    smtp.Send(correo);
                    return true;
                } catch (System.Net.Mail.SmtpException ex) {
                    String res = ex.Message;
                    return false;
                }
            } else
                return false;
        }


    }
}





