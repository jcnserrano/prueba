using Correos.SimuladorOfertas.Common;
using System;
using System.IO;
namespace Correos.SimuladorOfertas.Business
{
    public class GestorFicherosLogBL
    {
        /// <summary>
        /// Realiza la operación atómica de enviar un mail con los 
        /// </summary>
        /// <param name="conLog">Indica si se debe adjuntar el fichero del log</param>
        /// <param name="conXML">Indica si se debe adjuntar el fichero xml de la oferta sincronizada</param>
        /// <param name="oferta">Contiene el guid de la oferta para recuperar el fichero del xml generado</param>        
        public void EnviarCorreoFicheros(bool conLog, bool conXML, Guid oferta)
        {
            //Genera una nueva instancia para enviar el correo
            try
            {
                #region version anterior
                //var oApp = new Microsoft.Office.Interop.Outlook.Application();                

                //Microsoft.Office.Interop.Outlook.NameSpace ns = oApp.GetNamespace("MAPI");
                //var f = ns.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);

                //System.Threading.Thread.Sleep(1000);

                //var mailItem = (Microsoft.Office.Interop.Outlook.MailItem)oApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);

                //mailItem.Subject = SimuladorResources.MailSubject;
                //mailItem.HTMLBody = SimuladorResources.MailHTMLBody;
                //mailItem.To = SimuladorResources.MailTo;
                //string ficheroLog = System.IO.Path.Combine(Utils.ObtenerDirectorioBase(), "SimuladorLogUI.txt");
                //string ficheroXML = System.IO.Path.Combine(Utils.ObtenerDirectorioBase(), oferta + "_sync.xml");

                //if ((conLog) && (File.Exists(ficheroLog)))
                //{
                //    mailItem.Attachments.Add(ficheroLog, Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue);
                //}
                //if ((conXML) && (File.Exists(ficheroXML)))
                //{
                //    mailItem.Attachments.Add(ficheroXML, Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue);
                //}               

                //mailItem.Send();
                //MessageBox.Show(SimuladorResources.ExitoEnvioFicheroErrores, SimuladorResources.MensajeExito, MessageBoxButtons.OK, MessageBoxIcon.Information, 0);
                #endregion

                AyudaEnviarCorreo mapi = new AyudaEnviarCorreo();

                string ficheroLog = System.IO.Path.Combine(Utils.ObtenerDirectorioBase(), "SimuladorLogUI.txt");
                string ficheroXML = System.IO.Path.Combine(Utils.ObtenerRutaSubidaXML(), oferta + ".xml");

                if ((conLog) && (File.Exists(ficheroLog)))
                {
                    mapi.AddAttachment(ficheroLog);
                }
                if ((conXML) && (File.Exists(ficheroXML)))
                {
                    mapi.AddAttachment(ficheroXML);
                }

                mapi.AddRecipientTo(SimuladorResources.MailTo);
                mapi.SendMailPopup(SimuladorResources.MailSubject, SimuladorResources.MailHTMLBody);
            }
            catch (Exception ex)
            {
                RegistrarAccionesSimulador.GuardarExcepcion(ex);
            }
        }

    }
}
