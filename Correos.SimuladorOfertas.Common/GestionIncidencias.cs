using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Correos.SimuladorOfertas.Common
{
    public static class GestionIncidencias
    {
        #region Reportar Errores Usuarios

        public static void ReportarIncidencia(String usuario, String password)
        {
            bool tieneAccesoDirectorio;
            String pathDirectorio = string.Format(Utils.GetValorFromAppConfig(AppSettingsEnum.RutaIncidencias), String.Empty);

            try
            {
                tieneAccesoDirectorio = Directory.Exists(pathDirectorio);
            }
            //catch (UnauthorizedAccessException)
            catch
            {
                tieneAccesoDirectorio = false;
            }

            //Realizamos la copia de los ficheros     
            if (tieneAccesoDirectorio)
            {
                CopiaArchivosLocalesUsuario(usuario);
                MessageBox.Show("Se ha terminado de enviar la información.", "Fin del envío", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                if (password != null)
                {
                    using (new Impersonation("CORREOS", usuario, password))
                    {
                        CopiaArchivosLocalesUsuario(usuario);
                        MessageBox.Show("Se ha terminado de enviar la información.", "Fin del envío", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Por favor, conéctese con SAP CRM para enviar la información.", "Fin del envío", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }

            //MessageBox.Show("Se ha terminado de enviar la información.", "Fin del envío", MessageBoxButtons.OK, MessageBoxIcon.Information);

            //Envíamos la notificación de la incidencia por correo electrónico
            //DialogResult confirmResult = MessageBox.Show("¿Desea notificar la incidencia a Web Quorum?", "Incidencia Generador.", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            //if (confirmResult == DialogResult.Yes)
            //{
            //    NotificarIncidenciaPorMail(usuario);
            //}
        }

        private static void CopiaArchivosLocalesUsuario(String usuario)
        {
            try
            {
                //Obtenemos las rutas
                String pathBBDD_Local = System.IO.Path.Combine(Utils.ObtenerDirectorioBase(), "Simulador.sdf");
                String pathLog_Local = System.IO.Path.Combine(Utils.ObtenerDirectorioBase(), "SimuladorLogUI.txt");
                String pathIncidencias = string.Format(Utils.GetValorFromAppConfig(AppSettingsEnum.RutaIncidencias), usuario);

                /* Una vez conectados comprobamos en Inc_ProCom si existe una carpeta para el usuario. Si no existe, la creamos.*/
                if (!System.IO.Directory.Exists(pathIncidencias))
                {
                    System.IO.Directory.CreateDirectory(pathIncidencias);
                }

                //Si dentro de la carpeta del usuario ya hay ficheros, los copiamos en una nueva carpeta backup
                //añadiendoles a su nombre la fecha y hora actuales. (Ej: SimuladorLogUI_2202171030.txt).

                //Path de los archivos "Simulador.sdf" y "SimuladorLogUI.txt" dentro de la carpeta Inc_Procom 

                DateTime date = DateTime.Now;
                string fechaEnvio = date.ToString("yyyy-MM-dd_hhmm");



                String pathBBDD_Server = System.IO.Path.Combine(pathIncidencias, "Simulador.sdf");
                String pathLog_Server = System.IO.Path.Combine(pathIncidencias, "SimuladorLogUI.txt");

                //Creo el path de la carpeta Backup
                string pathCarpetaBackup = System.IO.Path.Combine(pathIncidencias + "\\Backup\\");
                string fechaStr = DateTime.Now.ToString("ddmmyyyy_hhmm");
                string pathBBDD_Backup = System.IO.Path.Combine(pathCarpetaBackup + "Simulador_" + fechaStr + ".sdf");
                string pathLog_Backup = System.IO.Path.Combine(pathCarpetaBackup + "SimuladorLogUI_" + fechaStr + ".txt");

                if (File.Exists(pathBBDD_Server) && File.Exists(pathLog_Server))
                {
                    //Compruebo que exista la carpeta y, si no existe, la creo
                    if (!Directory.Exists(pathCarpetaBackup))
                    {
                        Directory.CreateDirectory(pathCarpetaBackup);
                    }

                    //Copiamos .sdf y log.txt que están dentro de la carpeta del usuario en Inc_Procom a la subcarpeta BACKUP
                    System.IO.File.Copy(pathBBDD_Server, pathBBDD_Backup, true);
                    System.IO.File.Copy(pathLog_Server, pathLog_Backup, true);
                }

                //Copiamos .sdf y logui.txt de local a la carpeta del usuario en Inc_Procom
                System.IO.File.Copy(pathBBDD_Local, pathBBDD_Server, true);
                System.IO.File.Copy(pathLog_Local, pathLog_Server, true);


                CopiaXmlOfertasUsuario(usuario);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Se ha producido un error en la generación de la incidencia. Compruebe que su equipo está conectado a la red y vuelva a intentarlo", "Incidencia Generador.", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        //-----------------------------------------------------------------
        // ARCHIVOS .GZ SOLO SE PUEDE COMPRIMIR CON ESTO EN FRAMEWORK 4.0
        // CUANDO SE CAMBIE A FRAMEWORK 4.5
        //-----------------------------------------------------------------
        public static void CopiaXmlOfertasUsuario(String usuario)
        {
            String pathXmlSubida = Utils.ObtenerRutaSubidaXML();

            //cunado lleve un tiempo instalado se podrá poner "ZCCrearOportOferta_.*" EN LUGAR DE "*-*"
            ZipearXmlOfertasUsuario(usuario, pathXmlSubida, "ofertas_subida.gz", "*-*");


            String pathXmlDescarga = Utils.ObtenerRutaDescargaXML();

            ZipearXmlOfertasUsuario(usuario, pathXmlDescarga, "ofertas_descarga_fase1.gz", "*_ZCOfertasGestorRfc.xml");
            ZipearXmlOfertasUsuario(usuario, pathXmlDescarga, "ofertas_descarga_fase2.gz", "*_ReDescargaSAP.xml");
            ZipearXmlOfertasUsuario(usuario, pathXmlDescarga, "ofertas_descarga_response.gz", "ZCCrearOportOferta_Response*");
            ZipearXmlOfertasUsuario(usuario, pathXmlDescarga, "clientes_descarga_response.gz", "ZCClientesGestorRfc_*");

            //ZCCrearOportOferta_Response
        }


        //-------------------------------------------------------------
        // ARCHIVOS .ZIP
        // CUANDO SE CAMBIE A FRAMEWORK 4.5
        //-------------------------------------------------------------
        public static void CopiaXmlOfertasUsuario_ZIP(String usuario)
        {
            //String pathXmlSubida = Utils.ObtenerRutaSubidaXML();

            //ZipearXmlOfertasUsuario_ZIP(usuario, pathXmlSubida, "ofertas_subida.zip", "*-*");


            //String pathXmlDescarga = Utils.ObtenerRutaDescargaXML();

            //ZipearXmlOfertasUsuario_ZIP(usuario, pathXmlDescarga, "ofertas_descarga_fase1.zip", "*_ZCOfertasGestorRfc.xml");
            //ZipearXmlOfertasUsuario_ZIP(usuario, pathXmlDescarga, "ofertas_descarga_fase2.zip", "*_ReDescargaSAP.xml");
            //ZipearXmlOfertasUsuario_ZIP(usuario, pathXmlDescarga, "ofertas_descarga_response.zip", "ZCCrearOportOferta_Response*");


        }


        public static void ZipearXmlOfertasUsuario(String usuario, string pathXml, string nombreZip, string buscar)
        {
            int ind = 0;

            List<string> listaFicheros = new List<string>();

            DateTime Desde = DateTime.Now.AddMonths(-2);
            DateTime Hasta = DateTime.Now;

            DirectoryInfo dirInfo = new DirectoryInfo(pathXml);
            

            try
            {
                foreach (var fich in dirInfo.GetFiles(buscar))
                {
                    DateTime creacion = File.GetCreationTime(fich.FullName);
                    DateTime modificacion = File.GetLastWriteTime(fich.FullName);


                    if (modificacion >= Desde && modificacion <= Hasta)
                    {
                        Debug.WriteLine(fich.Name + " - " + creacion + " - " + modificacion);
                        listaFicheros.Add(fich.FullName);
                        
                    }
                }

                if (listaFicheros.Count > 0)
                {

                    string ficheroZip = System.IO.Path.Combine(pathXml, nombreZip);

                    if (File.Exists(ficheroZip))
                    {
                        File.Delete(ficheroZip);
                    }

                    CompressListFile(pathXml, ficheroZip, listaFicheros);

                    String pathIncidencias = string.Format(Utils.GetValorFromAppConfig(AppSettingsEnum.RutaIncidencias), usuario);
                    String pathIncidenciasServer = System.IO.Path.Combine(pathIncidencias, nombreZip);


                    System.IO.File.Copy(ficheroZip, pathIncidenciasServer, true);

                }
            }
            catch (System.Exception ex)
            {
            }

        }

        static void CompressListFile(string sInDir, string sOutFile, List<string> lstInFiles)
        {
            //PARA DESCOMPRIMIR. CON LA FUNCIÓN DecompressToDirectory

            string[] sFiles = lstInFiles.ToArray();

            int iDirLen = sInDir[sInDir.Length - 1] == Path.DirectorySeparatorChar ? sInDir.Length : sInDir.Length + 1;

            using (FileStream outFile = new FileStream(sOutFile, FileMode.Create, FileAccess.Write, FileShare.None))
            using (GZipStream str = new GZipStream(outFile, CompressionMode.Compress))
                foreach (string sFilePath in sFiles)
                {
                    string sRelativePath = sFilePath.Substring(iDirLen);
                    CompressFile(sInDir, sRelativePath, str);
                }
        }
        static void CompressFile(string sDir, string sRelativePath, GZipStream zipStream)
        {
            //Compress file name
            char[] chars = sRelativePath.ToCharArray();
            zipStream.Write(BitConverter.GetBytes(chars.Length), 0, sizeof(int));
            foreach (char c in chars)
                zipStream.Write(BitConverter.GetBytes(c), 0, sizeof(char));

            //Compress file content
            byte[] bytes = File.ReadAllBytes(Path.Combine(sDir, sRelativePath));
            zipStream.Write(BitConverter.GetBytes(bytes.Length), 0, sizeof(int));
            zipStream.Write(bytes, 0, bytes.Length);
        }

        static void DecompressToDirectory(string sCompressedFile, string sDir)
        {
            using (FileStream inFile = new FileStream(sCompressedFile, FileMode.Open, FileAccess.Read, FileShare.None))
            using (GZipStream zipStream = new GZipStream(inFile, CompressionMode.Decompress, true))
                while (DecompressFile(sDir, zipStream)) ;
        }

        static bool DecompressFile(string sDir, GZipStream zipStream)
        {
            //Decompress file name
            byte[] bytes = new byte[sizeof(int)];
            int Readed = zipStream.Read(bytes, 0, sizeof(int));
            if (Readed < sizeof(int))
                return false;

            int iNameLen = BitConverter.ToInt32(bytes, 0);
            bytes = new byte[sizeof(char)];
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < iNameLen; i++)
            {
                zipStream.Read(bytes, 0, sizeof(char));
                char c = BitConverter.ToChar(bytes, 0);
                sb.Append(c);
            }
            string sFileName = sb.ToString();

            //Decompress file content
            bytes = new byte[sizeof(int)];
            zipStream.Read(bytes, 0, sizeof(int));
            int iFileLen = BitConverter.ToInt32(bytes, 0);

            bytes = new byte[iFileLen];
            zipStream.Read(bytes, 0, bytes.Length);

            string sFilePath = Path.Combine(sDir, sFileName);
            string sFinalDir = Path.GetDirectoryName(sFilePath);
            if (!Directory.Exists(sFinalDir))
                Directory.CreateDirectory(sFinalDir);

            using (FileStream outFile = new FileStream(sFilePath, FileMode.Create, FileAccess.Write, FileShare.None))
                outFile.Write(bytes, 0, iFileLen);

            return true;
        }
        //-------------------------------------------------------------
        // ARCHIVOS .ZIP
        // CUANDO SE CAMBIE A FRAMEWORK 4.5
        //-------------------------------------------------------------
        public static void ZipearXmlOfertasUsuario_ZIP(String usuario, string pathXml, string nombreZip, string buscar)
        {
            //List<string> listaFicheros = new List<string>();

            //DateTime Desde = DateTime.Now.AddMonths(-2);
            //DateTime Hasta = DateTime.Now;

            //DirectoryInfo dirInfo = new DirectoryInfo(pathXml);

            //try
            //{
            //    foreach (var fich in dirInfo.GetFiles(buscar))
            //    {
            //        DateTime creacion = File.GetCreationTime(fich.FullName);
            //        DateTime modificacion = File.GetLastWriteTime(fich.FullName);

            //        if (modificacion >= Desde && modificacion <= Hasta)
            //        {
            //            Debug.WriteLine(fich.Name + " - " + creacion + " - " + modificacion);
            //            listaFicheros.Add(fich.FullName);
            //        }
            //    }

            //    if (listaFicheros.Count > 0)
            //    {
            //        string ficheroZip = System.IO.Path.Combine(pathXml, nombreZip);

            //        if (File.Exists(ficheroZip))
            //        {
            //            File.Delete(ficheroZip);
            //        }

            //        using (ZipArchive archive = ZipFile.Open(ficheroZip, ZipArchiveMode.Create))
            //        {
            //            foreach (string file in listaFicheros)
            //            {
            //                archive.CreateEntryFromFile(file, Path.GetFileName(file), CompressionLevel.Optimal);
            //            }
            //        }

            //        String pathIncidencias = string.Format(Utils.GetValorFromAppConfig(AppSettingsEnum.RutaIncidencias), usuario);
            //        String pathIncidenciasServer = System.IO.Path.Combine(pathIncidencias, nombreZip);


            //        System.IO.File.Copy(ficheroZip, pathIncidenciasServer, true);

            //    }
            //}
            //catch (System.Exception ex)
            //{
            //}

        }

        public static void NotificarIncidenciaPorMail(String usuario)
        {
            try
            {
                //Abrimos el correo para envíarlo
                DialogResult editarCorreo = MessageBox.Show("¿Desea editar el correo antes de envíarlo?.", "Incidencia Generador.", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                //Creamos el mensaje personalizado
                String mensaje = "<P>{0}, <BR></BR><BR></BR>" +
                                    "Se ha producido un error en Generador de Ofertas. Por favor, notifiquen del mismo al equipo de resolución de incidencias." +
                                    "<BR></BR><BR></BR>" +
                                    "Muchas gracias," +
                                    "<BR></BR><BR></BR>" +
                                    "Un saludo.</P>";

                //Preparamos el saludo y la ruta del servidor incprocom
                String saludo = (System.DateTime.Now.Hour < 14) ? "Buenos días" : "Buenas tardes";

                mensaje = string.Format(mensaje, saludo);

                //Creamos el correo en outlook
                var appOutlook = new Microsoft.Office.Interop.Outlook.Application();

                MailItem mailItem = (MailItem)appOutlook.CreateItem(OlItemType.olMailItem);

                //Indicamos el asunto y destinatario
                mailItem.Subject = string.Format("Resolución de Incidencia Generador. Usuario {0}.", usuario);
                mailItem.To = "mmunozarchidona@deloitte.es";

                //Añadimos la firma y el mensaje
                mailItem.GetInspector.Activate();
                var signature = mailItem.HTMLBody;
                mailItem.HTMLBody = mensaje + signature;

                if (editarCorreo == DialogResult.Yes)
                {
                    mailItem.Display(false);
                }
                else
                {
                    mailItem.Send();
                    MessageBox.Show("Se ha envíado el correo de notificación a QUORUM.", "Incidencia Generador.", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Se ha producido un error en el envío/edición del correo electrónico. Notifique la incidencia a Quorum envíando un correo a la dirección quorumw@correos.com.",
                                "Incidencia Generador.", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
        #endregion
}
