using Correos.SimuladorOfertas.Common;
using Correos.SimuladorOfertas.Common.Enums;
using Correos.SimuladorOfertas.Common.Extensions;
using Correos.SimuladorOfertas.DTOs;
using Correos.SimuladorOfertas.Persistence;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace Correos.SimuladorOfertas.Business
{
    public class InformesBL
    {
        #region Métodos Publicos

        #region GenerarInformeDetrec
        /// <summary>
        /// Genera el Informe Detrec para la Evalualción de Servicio de Recogida
        /// </summary>
        /// <param name="FicheroOrigen"></param>
        /// <param name="FicheroDestino"></param>
        /// <param name="cliente"></param>
        /// <param name="oferta"></param>
        public bool GenerarInformeDetrec(string FicheroDestino, OfertaBE oferta, ClienteBE cliente)
        {
            bool exito = true;

            //Obtenemos el dichero plantilla
            string FicheroOrigen = string.Format(Utils.GetValorFromAppConfig(AppSettingsEnum.RutaPlantillaFichaDetrec), AppDomain.CurrentDomain.BaseDirectory);

            //FicheroDestino que se modificará para generar el informe
            ManagerExcel objExcel = new ManagerExcel(FicheroDestino, false);

            try
            {
                Dictionary<string, object> objContenidoARellenar = new Dictionary<string, object>();

                // se copia la Plantilla que es el FicheroOrigen hacia el destino seleccionado por el usuario
                File.Copy(FicheroOrigen, FicheroDestino, true);

                // Se rellena el objeto con todas las celdas a escribir
                objContenidoARellenar.Add("E9", cliente.Provincia);
                //DateTime dt = Convert.ToDateTime(oferta.FechaCreacion);
                //objContenidoARellenar.Add("R9", dt.ToShortDateString());                
                objContenidoARellenar.Add("R11", cliente.CodClienteSAP);
                objContenidoARellenar.Add("J13", cliente.Nombre);
                objContenidoARellenar.Add("E15", oferta.PersonaContacto);
                objContenidoARellenar.Add("S15", cliente.Telefono);

                objExcel.AbrirFichero();
                if (objExcel.Abierto)
                {
                    objExcel.SeleccionarHoja(SimuladorResources.FichaDetrec);
                    objExcel.EscribirCeldasMultiple(objContenidoARellenar);
                    objExcel.GuardarLibro();
                    objExcel.CerrarExcel();
                }
            }
            catch (Exception ex)
            {
                RegistrarAccionesSimulador.GuardarExcepcion(ex);
                exito = false;
            }
            finally
            {
                if (objExcel.Abierto)
                {
                    objExcel.CerrarExcel();
                }
            }
            return exito;

        }

        #endregion

        #region GenerarFichaCliente

        /// <summary>
        /// Función que genera la FichaCliente en la ruta donde que se le pasa
        /// </summary>
        /// <param name="FicheroOrigen"></param>
        /// <param name="FicheroDestino"></param>
        /// <param name="Oferta"></param>
        /// <param name="Cliente"></param>
        public bool GenerarFichaCliente(String FicheroDestino, OfertaBE ofertaFicha, ClienteBE clienteFicha)
        {
            bool exito = true;

            //Obtenemos el dichero plantilla
            string FicheroOrigen = string.Format(Utils.GetValorFromAppConfig(AppSettingsEnum.RutaPlantillaFichaCliente), AppDomain.CurrentDomain.BaseDirectory);


            //creamos el manager de exel. (aún no hay ninguna instancia abierta ni nada).
            ManagerExcel objExcel = new ManagerExcel(FicheroDestino, false);

            try
            {
                //Creamos la variable que contendra las celdas y sus valores.
                Dictionary<string, object> objContenidoARellenar = new Dictionary<string, object>();
                //Copiamos el fichero del origen al destino.
                File.Copy(FicheroOrigen, FicheroDestino, true);

                //se rellenan las celdas que se desean insertar el valor
                //objContenidoARellenar.Add("D5", clienteFicha.Nombre);
                //objContenidoARellenar.Add("D7", clienteFicha.CIF);
                //objContenidoARellenar.Add("D8", clienteFicha.Direccion);
                //objContenidoARellenar.Add("D9", clienteFicha.Ciudad);
                //objContenidoARellenar.Add("D10", clienteFicha.CP);
                //ProductoOfertaBL objProductoOferta = new ProductoOfertaBL();

                //se rellenan las celdas que se desean insertar el valor
                objContenidoARellenar.Add("D28", clienteFicha.Nombre);
                objContenidoARellenar.Add("D30", clienteFicha.CIF);
                objContenidoARellenar.Add("D31", clienteFicha.Direccion);
                objContenidoARellenar.Add("D32", clienteFicha.Ciudad);
                objContenidoARellenar.Add("D33", clienteFicha.CP);
                objContenidoARellenar.Add("D35", clienteFicha.Telefono);                
                ProductoOfertaBL objProductoOferta = new ProductoOfertaBL();

                objContenidoARellenar.Add("H29", objProductoOferta.ObtenerSumatorioNumerosEnviosOferta(ofertaFicha.idOferta));

                //Una vez copiado podemos abrirlo. (en este punto, al aejecutar la instruccion, ya hay una tarea abierda de excel para el generar el infore).
                objExcel.AbrirFichero();
                if (objExcel.Abierto)
                {
                    //Si lo ha podido abrir entonces Seleccionamos la hoja que deseamos modificar.
                    objExcel.SeleccionarHoja(SimuladorResources.FichaCliente);
                    //Rellenamos las celdas en la hoja que hemos seleccionado.
                    objExcel.EscribirCeldasMultiple(objContenidoARellenar);
                    //Una vez que terminamos de generar el excel guardamos cambios y cerramos.
                    objExcel.GuardarLibro();
                    objExcel.CerrarExcel();
                }

            }
            catch (Exception ex)
            {
                RegistrarAccionesSimulador.GuardarExcepcion(ex);
                exito = false;
            }
            finally
            {
                //por si acaso ha habido un error y se queda la instancia abierta del excel se cierra, (Evitamos perdida de rendimiento del procesador).
                if (objExcel.Abierto)
                {
                    objExcel.CerrarExcel();
                }
            }
            return exito;
        }

        #endregion

        #region GenerarInformes

        #region InformesResumen
        /// <summary>
        /// Método generación de informes resumen de precios y tarifas
        /// </summary>
        /// <param name="objProductoBE">productoBE del que se muestra el informe</param>
        /// <param name="objProductoOfertaBE">productoOfertaBE del que se muestra el informe</param>
        /// <param name="fechaInicial">fecha inicial de validez de la oferta</param>
        /// <param name="fechaFinal">fecha final de validez de la oferta</param>
        /// <param name="nombreCliente">Nombre del cliente al que pertenece la oferta</param>
        /// <param name="tarifasConDescuento">indica si queremos el informe de precios o el de tarifas</param>
        public void GenerarInformeResumen(Collection<ProductoBE> listaObjProductoBE, Collection<ProductoOfertaBE> listaObjProductoOfertaBE, string fechaInicial, string fechaFinal, string nombreCliente, bool tarifasConDescuento, TipoInforme tipoInforme)
        {
            #region Variables

            Collection<string> listadoFicherosTemporales = new Collection<string>();
            string ficheroFinal = string.Empty;
            string plantillaNombres, nombreInformeResumen, nombreInforme;
            int numInformesTotales = listaObjProductoBE.Where(x => x.DebeGenerarInforme).Count();

            // Utilizada para los nombres de los ficheros
            int i = 0;

            plantillaNombres = "{0}_{1}." + ((tipoInforme == TipoInforme.DOCX) ? "docx" : "xlsx");
            nombreInformeResumen = (tarifasConDescuento) ? "InformeResumenPrecios" : "InformeResumenTarifas";
            nombreInforme = (tarifasConDescuento) ? "InformePrecios" : "InformeTarifas";
            ficheroFinal = Path.Combine(System.IO.Path.GetTempPath(), string.Format(CultureInfo.InvariantCulture, plantillaNombres, nombreInformeResumen, DateTime.Now.ToString("yyyyMMddhhmmss")));
            
            #endregion

            #region Bucle
                                    
            foreach (ProductoOfertaBE productoOferta in listaObjProductoOfertaBE)
            {
                string ficheroTemporal = string.Empty;

                //Comprobamos si es el último de los informes                
                bool esUltimoInforme = (i == numInformesTotales - 1);
                                
                // Obtenemos el producto que corresponde con el ProductoOferta
                ProductoBE producto = listaObjProductoBE.Where(x => x.idProducto == productoOferta.idProducto).FirstOrDefault();

                if (!producto.DebeGenerarInforme)
                {
                    continue;
                }

                //Generamos el nombre del fichero temporal
                ficheroTemporal = Path.Combine(System.IO.Path.GetTempPath(), string.Format(CultureInfo.InvariantCulture, plantillaNombres, nombreInforme + producto.CodProducto + productoOferta.CodModalidadNegociacion, DateTime.Now.ToString("yyyyMMddhhmmss")));

                    // Añadimos los ficheros al array
                listadoFicherosTemporales.Add(ficheroTemporal);

                // Llamada al flujo habitual del programa con los boolean señalando que es un informe múltiple
                this.GenerarInforme(producto, productoOferta, fechaInicial, fechaFinal, nombreCliente, tarifasConDescuento, true, ficheroTemporal, ficheroTemporal, tipoInforme, esUltimoInforme);

                i++;
            }

            //Si es Word, en formato docx
            if (tipoInforme == TipoInforme.DOCX)
            {
                this.generarDocx(listadoFicherosTemporales, ficheroFinal);

                foreach (string fichero in listadoFicherosTemporales)
                {
                    if (File.Exists(fichero))                    
                    File.Delete(fichero);
                }
            }
            //Si es Excel en formato xlsx
            else
            {
                ManagerExcel.CombineWorkBooks(ficheroFinal, listadoFicherosTemporales.ToArray(), true);
                ManagerExcel.AbrirExcelStandalone(ficheroFinal);                
            }
            #endregion
        }


        /// <summary>
        /// Función que concatena todos los PDFs enviados por parámetro
        /// </summary>
        /// <param name="ficherosPdf">Collection de ficheros PDF</param>
        /// <param name="ficheroFinalPdf">Fichero PDF donde se guardará el resultado</param>
        private void generarPDF(Collection<string> ficherosPdf, string ficheroFinalPdf)
        {
            // Creación del objeto documento
            Document document = new Document();

            // Creación del writer
            PdfCopy writer = new PdfCopy(document, new FileStream(ficheroFinalPdf, FileMode.Create));
            if (writer == null)
            {
                return;
            }

            // Apertura del documento
            document.Open();

            // Para cada fichero, creamos un reader y copiamos las páginas usando el writer
            foreach (string fichero in ficherosPdf)
            {
                // Creación del reader
                PdfReader reader = new PdfReader(fichero);
                reader.ConsolidateNamedDestinations();

                // Adición del contenido recorriendo las páginas
                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    PdfImportedPage page = writer.GetImportedPage(reader, i);
                    writer.AddPage(page);
                }
                // Cerramos el reader
                reader.Close();
            }

            // Cerramos los objetos writer y document
            writer.Close();
            document.Close();

            foreach (string fichero in ficherosPdf)
            {
                if (File.Exists(fichero))
                {
                    File.Delete(fichero);
                }
            }

            // Abrimos el pdf
            if (System.IO.File.Exists(ficheroFinalPdf))
            {
                System.Diagnostics.Process proc = new System.Diagnostics.Process();
                proc.EnableRaisingEvents = false;
                proc.StartInfo.FileName = ficheroFinalPdf;
                proc.Start();
            }
        }


        /// <summary>
        /// Función que concatena todos los DOCXs enviados por parámetro
        /// </summary>
        /// <param name="ficherosDocx">Collection de ficheros DOCX</param>
        /// <param name="ficheroFinalDocx">Fichero DOCx donde se guardará el resultado</param>
        private void generarDocx(Collection<string> ficherosDocx, string ficheroFinalDocx)
        {           
            //// Concatenamos los docx
            //if(objWord.ConcatenarDocx(ficherosDocx))
            //{
            //    if (System.IO.File.Exists(ficheroFinalDocx))
            //    {
            //        System.Diagnostics.Process proc = new System.Diagnostics.Process();
            //        proc.EnableRaisingEvents = false;
            //        proc.StartInfo.FileName = ficheroFinalDocx;
            //        proc.Start();
            //    }
            //}

            //if (objWord.Abierto)
            //{
            //    objWord.CerrarWord();
            //}

            var objWord = new ManagerWord(ficheroFinalDocx, false);

            try
            {                
                objWord.ConcatenarDocx(ficherosDocx);
            }
            catch (Exception ex)
            {
                RegistrarAccionesSimulador.GuardarExcepcion(ex);
            }
            finally
            {
                //if (objWord.Abierto)
                //{
                //    objWord.CerrarWord();
                //}
                               
                //Una vez creado el fichero en temporal se abre.
                if (System.IO.File.Exists(ficheroFinalDocx))
                {
                    System.Diagnostics.Process proc = new System.Diagnostics.Process();
                    proc.EnableRaisingEvents = false;
                    proc.StartInfo.FileName = ficheroFinalDocx;
                    proc.Start();
                }
                
            }
        }

        #endregion

        #region InformesIndividuales
        /// <summary>
        /// Genera el informe de tarifas/precios en formato PDF
        /// </summary>
        /// <param name="objProductoBE">productoBE del que se muestra el informe</param>
        /// <param name="objProductoOfertaBE">productoOfertaBE del que se muestra el informe</param>
        /// <param name="fechaInicial">fecha inicial de validez de la oferta</param>
        /// <param name="fechaFinal">fecha final de validez de la oferta</param>
        /// <param name="nombreCliente">Nombre del cliente al que pertenece la oferta</param>
        /// <param name="tarifasConDescuento">indica si queremos el informe de precios o el de tarifas</param>
        /// <param name="informeMultiple">Indica si se debe crear un informe orientado a la generación de uno múltiple (sin guardar PDF y guardando Words)</param>
        /// <param name="esUltimoInforme">indica si es el último producto de un informe múltimple</param>
        public void GenerarInforme(ProductoBE objProductoBE, ProductoOfertaBE objProductoOfertaBE, string fechaInicial, string fechaFinal, string nombreCliente, bool tarifasConDescuento, bool informeMultiple, string ficheroTemporalWord, string ficheroTemporalPdf, TipoInforme tipoInforme, bool esUltimoInforme)
        {
            ////Miramos si es de paqueteria            
            List<TramoInformeBE> listaTramosInfTarifas = new TramoInformeBL().ObtenerTramosInformeProducto(objProductoBE.CodProducto);

            //Paqueteria kg, si no en g. Miramos si su modelo de descuento es paquetería, tramos, o aparece como paquete en el doc. de tarifas                        
            //Si es Publicorreo óptimo, no lo consideramos como paquetería     
            bool esPaqueteria = !ModeloDescuentoEnum.GetPaqueteriasQueSonPublicorreos().Contains(objProductoBE.CodProducto) && objProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Paqueteria)) ||
                              objProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Tramos)) ||
                              listaTramosInfTarifas.Count > 0;

            bool esPublicorreo = (objProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Publicorreo)) || 
                                    ModeloDescuentoEnum.GetPaqueteriasQueSonPublicorreos().Contains(objProductoBE.CodProducto));
            bool esGrupoTramo = (objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoGrupoTramo)) || 
                                    (objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorPrecioCiertoGrupoTramo));


            //Si es un informe de tarifas, y el producto es de paquetería, imprimimos el informe de descuento por destino en todos los casos
            if (!tarifasConDescuento && esPaqueteria)
                {
                //Generar Informe modelo standar
                if (tipoInforme == TipoInforme.DOCX)
                    {
                    this.GenerarInformeEstandar(objProductoBE, objProductoOfertaBE, fechaInicial, fechaFinal, nombreCliente, tarifasConDescuento, informeMultiple, ficheroTemporalWord, ficheroTemporalPdf, esUltimoInforme);
                    }
                    else if (tipoInforme == TipoInforme.EXCEL)
                    {
                    this.GenerarInformeEstandarExcel(objProductoBE, objProductoOfertaBE, fechaInicial, fechaFinal, nombreCliente, tarifasConDescuento, informeMultiple, ficheroTemporalWord);
                    }
                }
                else
                {
                if(esGrupoTramo)
                    {
                    //Generar Informe según modelo de grupos de tramos
                    if (tipoInforme == TipoInforme.DOCX)
                    {
                        this.GenerarInformeModeloGrupoTramos(objProductoBE, objProductoOfertaBE, fechaInicial, fechaFinal, nombreCliente, tarifasConDescuento, informeMultiple, ficheroTemporalWord, ficheroTemporalPdf, esUltimoInforme);
                    }
                    else if (tipoInforme == TipoInforme.EXCEL)
                    {
                        this.GenerarInformeModeloGrupoTramosExcel(objProductoBE, objProductoOfertaBE, fechaInicial, fechaFinal, nombreCliente, tarifasConDescuento, informeMultiple, ficheroTemporalWord);
                    }
                }
                else  if (!esGrupoTramo && esPublicorreo && (objProductoBE.Destinos.Where(x => x.Tramos.Count > 0).ToList().Count.Equals(3)))
            {
                    //Si es publicorreo y sólo tiene 3 destinos...                  
                    if (tipoInforme == TipoInforme.DOCX)
                {
                        this.GenerarInformeEstandarPubliCorreo(objProductoBE, objProductoOfertaBE, fechaInicial, fechaFinal, nombreCliente, tarifasConDescuento, informeMultiple, ficheroTemporalWord, ficheroTemporalPdf, esUltimoInforme);
                    }
                    else if (tipoInforme == TipoInforme.EXCEL)
                    {
                        this.GenerarInformeEstandarPubliCorreoExcel(objProductoBE, objProductoOfertaBE, fechaInicial, fechaFinal, nombreCliente, tarifasConDescuento, informeMultiple, ficheroTemporalWord);
                    }
                }
                else
                {
                    //Generar Informe modelo standar
                    if (tipoInforme == TipoInforme.DOCX)
                    {
                        this.GenerarInformeEstandar(objProductoBE, objProductoOfertaBE, fechaInicial, fechaFinal, nombreCliente, tarifasConDescuento, informeMultiple, ficheroTemporalWord, ficheroTemporalPdf, esUltimoInforme);
                    }
                    else if (tipoInforme == TipoInforme.EXCEL)
                    {
                        this.GenerarInformeEstandarExcel(objProductoBE, objProductoOfertaBE, fechaInicial, fechaFinal, nombreCliente, tarifasConDescuento, informeMultiple, ficheroTemporalWord);
                    }
                }
            }
        }

        /// <summary>
        /// Genera el informe de tarifas/precios en formato PDF
        /// </summary>
        /// <param name="objProductoBE">productoBE del que se muestra el informe</param>
        /// <param name="objProductoOfertaBE">productoOfertaBE del que se muestra el informe</param>
        /// <param name="fechaInicial">fecha inicial de validez de la oferta</param>
        /// <param name="fechaFinal">fecha final de validez de la oferta</param>
        /// <param name="nombreCliente">Nombre del cliente al que pertenece la oferta</param>
        /// <param name="tarifasConDescuento">indica si queremos el informe de precios o el de tarifas</param>
        /// <returns>Devuelve el nombre del fichero donde se ha guardado el fichero</returns>
        private void GenerarInformeModeloGrupoTramosExcel(ProductoBE objProductoBE, ProductoOfertaBE objProductoOfertaBE, string fechaInicial, string fechaFinal, string nombreCliente, bool tarifasConDescuento, bool informeMultiple, string ficheroTemporalExcel)
        {
            #region Variables

            //Rango de la primera celda de la tabla grupo de tramos
            Microsoft.Office.Interop.Excel.Range primeraCeldaGrupoTramos = null;
            Microsoft.Office.Interop.Excel.Range primeraCeldaTramos = null;

            //Nº de columnas requeridas para la tabla grupo de tramos
            int[] columnasRequeridasGrupoTramos = null; 

            //En caso de ser un producto con Destinos se rellena esta matriz de informacion
            object[,] matrizValores = null;
            object[] colsDescripcion;
            bool esTipoPrecioCierto = false;

            //Listado de valores que se ingresan en la tabla de valores añadidos si corresponde
            Collection<ReporteVABE> ListaTarifasVAReporte = new Collection<ReporteVABE>();

            //Lista de etiquetas con su correspondiente valor que se sustituye en el documento word
            Dictionary<string, string> objEtiquetas = new Dictionary<string, string>();

            //Se activa para los casos que se quiere mostrar la aclaración *Para expediciones.
            Boolean mostrarAclaracionExpediciones = false;

            //Contiene la ruta de la plantilla que se usa para generar el reporte de tarifas
            string rutaPlantilla = string.Format(Utils.GetValorFromAppConfig(AppSettingsEnum.RutaPlantillaInformeTarifas), AppDomain.CurrentDomain.BaseDirectory);
            rutaPlantilla = rutaPlantilla.Substring(0, rutaPlantilla.Length - 4) + "xlsx";

            string rutaLineaProducto = string.Format(Utils.GetValorFromAppConfig(AppSettingsEnum.RutaPlantillaLineaProducto), AppDomain.CurrentDomain.BaseDirectory, objProductoBE.PlantillaInformeTarifasPrecios);
            rutaLineaProducto = rutaLineaProducto.Substring(0, rutaLineaProducto.Length - 4) + "xlsx";

            string rutaPlantillaVA = string.Format(Utils.GetValorFromAppConfig(AppSettingsEnum.RutaPlantillaInformeTarifasVA), AppDomain.CurrentDomain.BaseDirectory);

            // Si no es un informe múltiple o no tiene nombre, poner nombre por defecto
            if (ficheroTemporalExcel.Equals(string.Empty) )//|| !informeMultiple)
            {
                if (tarifasConDescuento)
                {
                    ficheroTemporalExcel = Path.Combine(System.IO.Path.GetTempPath(), string.Format(CultureInfo.InvariantCulture, "{0}_{7}_{1}{2}{3}{4}{5}{6}.xlsx", "InformePrecios", DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(),
                        DateTime.Now.Day.ToString(), DateTime.Now.Hour.ToString(), DateTime.Now.Minute.ToString(), DateTime.Now.Second.ToString(), objProductoBE.CodProducto));
                }
                else
                {
                    ficheroTemporalExcel = Path.Combine(System.IO.Path.GetTempPath(), string.Format(CultureInfo.InvariantCulture, "{0}_{7}_{1}{2}{3}{4}{5}{6}.xlsx", "InformeTarifas", DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(),
                        DateTime.Now.Day.ToString(), DateTime.Now.Hour.ToString(), DateTime.Now.Minute.ToString(), DateTime.Now.Second.ToString(), objProductoBE.CodProducto));
                    //ficheroTemporalExcel = Path.Combine(System.IO.Path.GetTempPath(), string.Format(CultureInfo.InvariantCulture, "{0}_{1}.xlsx", "InformeTarifas", objProductoBE.CodProducto));
                }
            }

            //Ya tenemos generado el documento, ahora lo guardamos en PDF.                    
            string ficheroDestino = string.Empty;

            //instancia del objeto word con el que vamos  trabajar para generar el informe de tarifas
            ManagerExcel objExcel = new ManagerExcel(ficheroTemporalExcel, false);

            //Lista de tramos que usamos para, al generar el informe, mostrar la cabecera de los tramos.
            List<TramoBE> auxTramo = new List<TramoBE>();

            Collection<TramoBE> listaTramosEliminar = new Collection<TramoBE>();

            #endregion

            try
            {
                #region Obtencion Datos

                // Guardamos si es precio cierto o tipo descuento
                if (objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoPrecioCierto) || objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorPrecioCiertoGrupoTramo))
                {
                    esTipoPrecioCierto = true;
                }
                else
                {
                    esTipoPrecioCierto = false;
                }
                //Obtenemos los datos que ha ingresado el usuario.
                ConfiguracionGruposTramoBL configuracionGruposTramoBL = new ConfiguracionGruposTramoBL();
                ConfiguracionProductosBL configProducto = new ConfiguracionProductosBL();

                Collection<GrupoTramoBE> listaConfiguracionGrupoTramos = configuracionGruposTramoBL.ObtenerListaGruposTramoOferta(objProductoBE, objProductoOfertaBE.idProductoOferta);
                Collection<ConfiguracionTramoOfertaBE> listaConfiguracionTramo = configProducto.ObtenerConfiguracionTramoOferta(objProductoOfertaBE.idProductoOferta);
                Collection<InformacionDestinosBE> listaDestinos = new InformacionDestinosBL().ObtenerListadoInformacionDestinos(objProductoBE.CodProducto);

                //Se obtienen los destinos y tramos para los descuentos sobre las tarifas
                List<String> listaDestinosGT = listaConfiguracionGrupoTramos.OrderBy(t => t.OrdenDestino).GroupBy(t => t.CodDestino).Select(t => t.First()).Select(t => t.CodDestino).ToList();
                List<GrupoTramoBE> listaTramosGT = listaConfiguracionGrupoTramos.OrderBy(t => t.OrdenDestino).OrderBy(t => t.CodTramoInicialDecimal).GroupBy(t => t.Nombre).Select(t => t.First()).ToList();
                List<TramoInformeBE> listaTramosInfTarifas = new TramoInformeBL().ObtenerTramosInformeProducto(objProductoBE.CodProducto);
                Collection<ConfiguracionDestinoOfertaBE> listaConfiguracionDestino = configProducto.ObtenerConfiguracionDestinoOferta(objProductoOfertaBE.idProductoOferta);

                var listadoDestinos = objProductoBE.Destinos.Where(x => x.Tramos.Count > 0).OrderBy(t => t.Orden).ToList();

                if (objProductoBE.Internacional)
                {
                    listadoDestinos = (from d in listadoDestinos
                                       join c in listaConfiguracionDestino on d.idDestino equals c.idDestino
                                       where c.Distribucion.HasValue && c.Distribucion.Value > 0
                                       select d).ToList();

                }

                int numDestinos = listadoDestinos.Count();

                //Producto con destinos
                int numTramos = 0;
                //obtenemos el número máximo de tramos de entre los destinos. Ademas nos guardamos los tramos para la cabecera de fila de la tabla del informe. 
                foreach (DestinoBE destino in listadoDestinos) 
                {
                    if (numTramos < destino.Tramos.Count)
                    {
                        numTramos = destino.Tramos.Count;
                        TramoBE[] copiaTramos = new TramoBE[numTramos];
                        destino.Tramos.CopyTo(copiaTramos , 0);
                        auxTramo = copiaTramos.ToList();
                    }
                }

                //En este punto ya tenemos en auxTramo la colleccion de tramos que hay que mostrar.
                foreach (TramoBE item in auxTramo)
                {
                    //Eliminamos los tramos que no estén en ningun grupo de tramo
                    bool estaTramoEnGrupos = false;
                    foreach (GrupoTramoBE itemGrTamo in listaConfiguracionGrupoTramos)
                    {
                        if ((item.CodTramoDecimal >= itemGrTamo.CodTramoInicialDecimal) && (item.CodTramoDecimal <= itemGrTamo.CodTramoFinalDecimal))
                        {
                            estaTramoEnGrupos = true;
                            break;
                        }
                    }
                    if (!estaTramoEnGrupos)
                    {
                        listaTramosEliminar.Add(item);
                    }
                }

                if (auxTramo.Count == listaTramosEliminar.Count)
                {
                    listaTramosEliminar.Clear();
                }

                numTramos = auxTramo.Count;
                //Eliminamos los tramos que no deben insertarse
                foreach (TramoBE item in listaTramosEliminar)
                {
                    auxTramo.Remove(item);
                    numTramos--;
                }

                matrizValores = new object[numTramos, numDestinos];
                int i = 0;
                int j = 0;
                decimal maxDecimalesTarifa = 2;

                //Para los productos S0134 y S0235 y los destinos Z7, Z8 y Z9 No se deben mostrar los tramos de expediciones
                //JCNS. MOSTRAR TRAMO. LO CAMBIO DE SITIO
                //bool DebeMostrarTramo = true;


                foreach (DestinoBE destino in listadoDestinos)
                {
                    j = 0;

                    foreach (TramoBE tramo in destino.Tramos)
                    {
                        //Para los productos S0134, S0132 y S0235 y los destinos Z7, Z8 y Z9 No se deben mostrar los tramos de expediciones
                        //JCNS. MOSTRAR TRAMO. LO CAMBIO DE SITIO
                        bool DebeMostrarTramo = true;
                        if ((objProductoBE.CodProducto.Equals("S0236") || objProductoBE.CodProducto.Equals("S0133") ||objProductoBE.CodProducto.Equals("S0134") || objProductoBE.CodProducto.Equals("S0235") || objProductoBE.CodProducto.Equals("S0132")) && (destino.CodDestinoSAP.Equals("Z7") || destino.CodDestinoSAP.Equals("Z8") || destino.CodDestinoSAP.Equals("Z9")) && (tramo.CodTramo.StartsWith("E")))
                        {
                            DebeMostrarTramo = false;
                        }

                        TramoBE auxEliminar = listaTramosEliminar.FirstOrDefault(x => x.CodTramo.Equals(tramo.CodTramo));
                        if (auxEliminar == null)
                        {
                            TramoBE aux = auxTramo.FirstOrDefault(x => x.CodTramo.Equals(tramo.CodTramo));
                            if (aux != null)
                            {
                                object tarifaTramo = Math.Round(tramo.Tarifa, 5);

                                if (esTipoPrecioCierto)
                                {
                                    tarifaTramo = Math.Round(tramo.Tarifa, 5);
                                }

                                if (tarifasConDescuento)
                                {
                                    double auxTarifa = 0;

                                    //MMunoz FIX: No cogía el descuento del grupo, si no del tramo concreto. Lo hemos modificado, para que saque
                                    //el valor de Dto o PC del grupo.
                                    
                                    //Se obtiene la configuracion del grupo de tramos                                    
                                    GrupoTramoBE configGrupoTramo = listaConfiguracionGrupoTramos.FirstOrDefault(x => x.CodDestino.Equals(destino.CodDestinoSAP)
                                                        && x.CodTramoInicialDecimal <= tramo.CodTramoDecimal
                                                        && x.CodTramoFinalDecimal >= tramo.CodTramoDecimal
                                                        && x.idProductoOferta.Equals(objProductoOfertaBE.idProductoOferta));
                                                                                                            
                                    //Se obtiene la configuración del tramo
                                    ConfiguracionTramoOfertaBE configTramo = listaConfiguracionTramo.FirstOrDefault(x => x.idTramo.Equals(tramo.idTramo));
                                                                       
                                    if (configTramo != null)
                                    {                  
                                        //FIX: intentamos sacar el Dto o el PC del grupo. Si no hay, lo obtenemos de los valores del tramo.                               
                                        var descuentoFinal = (configGrupoTramo != null)? configGrupoTramo.DtoPC : configTramo.DescuentoFinal.Value;
                                        var precioCierto = (configGrupoTramo != null) ? configGrupoTramo.DtoPC : configTramo.PrecioCierto.Value;

                                        if ((objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoGrupoTramo)) &&
                                                configTramo.DescuentoFinal.HasValue && double.TryParse(descuentoFinal.ToString(), out auxTarifa))
                                        {
                                            auxTarifa = tramo.Tarifa - (tramo.Tarifa * auxTarifa / 100);
                                            tarifaTramo = Math.Round(auxTarifa, 5);
                                        }                                      
                                        else if ((objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorPrecioCiertoGrupoTramo)) &&
                                            configTramo.PrecioCierto.HasValue && double.TryParse(precioCierto.ToString(), out auxTarifa))
                                        {
                                            if (!auxTarifa.Equals(0))
                                            {
                                                tarifaTramo = Math.Round(auxTarifa, 5);
                                            }
                                        }
                                    }

                                }

                                if (DebeMostrarTramo)
                                {
                                    decimal argument = (decimal)(double)tarifaTramo;
                                    int count = BitConverter.GetBytes(decimal.GetBits(argument)[3])[2];

                                    if (count >= maxDecimalesTarifa)
                                        maxDecimalesTarifa = count;

                                    matrizValores[j, i] = tarifaTramo;
                                }
                                else
                                {
                                    matrizValores[j, i] = null;
                                }

                                j++;
                                DebeMostrarTramo = true;
                            }
                        }
                    }
                    i++;
                }

                //Se buscan los datos de los VA            
                ListaTarifasVAReporte = this.CrearTablaVA(objProductoBE.idProducto, objProductoOfertaBE.idProductoOferta, tarifasConDescuento);

                //Se busca el valor máximo de peso volumétrico
                int maxPesoVolumetrico; //40  o 60

                //Vemos si se ha definido un peso volumétrico máximo especifico para el tipo de producto actual (CodAnexoSAP)
                maxPesoVolumetrico = new ProductoSAPBL().ObtenerPesoVolumetricoMaxProducto(objProductoBE.CodProducto);

                //Si no hay valor máximo registrado, se coge el valor por defecto (40 Kg).
                if (maxPesoVolumetrico == 0) maxPesoVolumetrico = 40;
                String textoPesoVolumetrico = string.Format(SimuladorResources.PesoVolumetricoMaximoInforme, maxPesoVolumetrico);

                //Vemos si es de publicorreo
                Boolean esPublicorreo = objProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Publicorreo));

                #endregion

                #region Generacion Plantilla

                bool esNecesarioInsertarPlantilla = true;
                //if (!informeMultiple || !File.Exists(ficheroTemporalExcel))
                if (!File.Exists(ficheroTemporalExcel))
                {
                    esNecesarioInsertarPlantilla = false;
                    //Copiamos la plantilla en el fichero temporal.
                    File.Copy(rutaPlantilla, ficheroTemporalExcel, true);
                }

                //Abrimos el fichero
                try
                {
                    objExcel.AbrirFichero();
                }
                catch { }
                if (objExcel.Abierto)
                {
                    objExcel.SeleccionarPrimeraHojaLibreOCrearNuevaYRenombrar(objProductoBE.CodAnexoSAP + " - " + objProductoBE.CodProducto + " - " + GenerarAbreviaturaModeloNegociacion(objProductoOfertaBE.CodModalidadNegociacion.Trim()));

                    #region "Borrar una vez validado el método"

                    //if (!informeMultiple)
                    //{
                    //    objExcel.SeleccionarPrimeraHoja();
                    //}
                    //else
                    //{
                    //    objExcel.SeleccionarPrimeraHojaLibreOCrearNuevaYRenombrar(objProductoBE.CodAnexoSAP + " - " + objProductoBE.CodProducto);
                    //    if (esNecesarioInsertarPlantilla)
                    //    {
                    //        objExcel.InsertarDocumentoAlFinal(rutaPlantilla, 1);
                    //    }
                    //}

                    #endregion

                    #region Etiquetas excel

                    string textoParametro = string.Empty;
                    if (!tarifasConDescuento)
                    {
                        textoParametro = SimuladorResources.TituloReportTarifas;
                    }
                    else
                    {
                        textoParametro = SimuladorResources.TituloReportPrecios;
                    }

                    #endregion

                    #region Rellenar las tablas
                    int numTablasCrear = 1;
                    int numDestinosPorTabla = 15;

                    //Paqueteria kg, si no en g. Miramos si su modelo de descuento es paquetería, tramos, o aparece como paquete en el doc. de tarifas                        
                    //Si es Publicorreo óptimo, no lo consideramos como paquetería     
                    bool paqueteria = !ModeloDescuentoEnum.GetPaqueteriasQueSonPublicorreos().Contains(objProductoBE.CodProducto) && objProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Paqueteria)) ||
                                      objProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Tramos)) ||
                                      listaTramosInfTarifas.Count > 0;
                                        
                    //Si es de paquetería internacional
                    if (objProductoBE.Internacional)
                    {
                        numTablasCrear = (int)Math.Ceiling(numDestinos / 15.0);

                        if (numTablasCrear != 1)
                        {
                            numDestinosPorTabla = 15;
                        }
                        else
                        {
                            numDestinosPorTabla = numDestinos;
                        }
                    }
                    //Si es normal
                    else
                    {
                        numTablasCrear = 1;
                        numDestinosPorTabla = numDestinos;
                    }
                    
                    int numDestinosPorTablaInicial = numDestinosPorTabla;

                    //Creamos T tablas, de acuerdo al tipo de producto
                    for (int iTabla = 0; iTabla < numTablasCrear; iTabla++)
                    {
                        //Si es la última tabla, nos aseguramos del nº de columnas a mostrar    
                        if ((iTabla == numTablasCrear - 1) && (numDestinos % 15 != 0))
                        {
                            numDestinosPorTabla = numDestinos % 15;
                        }

                        if (listaConfiguracionGrupoTramos.Count > 0 && tarifasConDescuento)
                        {
                            objExcel.IniciarDibujarTabla(true, true);

                        string[] colsCabecera = new string[3];
                        colsCabecera[0] = "DESTINO";
                        colsCabecera[1] = "NOMBRE GRUPO TRAMO";

                        colsCabecera[2] = "PRECIO CIERTO";
                        string valorCaracter = "";
                        int redondeo = 5;

                        if (objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoGrupoTramo))
                        {
                            colsCabecera[2] = "DESCUENTO";
                            valorCaracter = "%";
                            redondeo = 2;
                        }

                            //Rellenar los datos de la tabla  
                            List<DestinoBE> destinos = listadoDestinos;
                            object[] colsGT = new object[numDestinosPorTabla + 1];

                            //añadir primera fila con las cabeceras de los destinos                                                     
                            colsDescripcion = new object[numDestinosPorTabla + 1];

                            //Añadimos la cabecera de los descuentos de grupo de tramos                            
                            if (!esTipoPrecioCierto)
                        {
                                colsDescripcion[0] = "Descuento aplicado sobre tarifa\t";
                                colsGT[0] = "GRUPOS DE TRAMOS\t";
                            }
                            else
                            {
                                colsDescripcion[0] = "GRUPOS DE TRAMOS\t";
                            }

                            for (int iDestino = 0; iDestino < numDestinosPorTabla; iDestino++)
                            {
                                DestinoBE objDestino = destinos[(iTabla * numDestinosPorTablaInicial) + iDestino];

                                //Obtenemos la descripcion
                                var objDescDestino = listaDestinos.FirstOrDefault(d => d.CodDestinoSAP.Equals(objDestino.CodDestinoSAPSinZona));
                                var descDestino = (objDescDestino != null && !String.IsNullOrEmpty(objDescDestino.DescripcionDestino)) ? objDescDestino.DescripcionDestino : "";
                                descDestino = StringUtil.ToTitleCaseIfAllUpper(descDestino);

                                if (objDescDestino != null)
                                {
                                    colsGT[iDestino + 1] = objDescDestino.CodDestinoSAP;
                                    colsDescripcion[iDestino + 1] = descDestino;
                                }
                                else
                                {
                                    colsGT[iDestino + 1] = objDestino.CodDestinoSAPSinZona;
                                    colsDescripcion[iDestino + 1] = String.Empty;
                                }
                            }

                            objExcel.AgregarFila(colsGT, obviarPrimeraColumna: true);
                            objExcel.AgregarFila(colsDescripcion, esAnchoFijo: true);

                            //Añadimos los descuentos para los grupos de tramos                        
                            for (int tr = 0; tr < listaTramosGT.Count; tr++)
                            {
                                GrupoTramoBE tramo = listaTramosGT[tr];
                                String descTramo = GetDescripionGrupoTramo(objProductoBE, tramo, listaTramosInfTarifas, listaTramosGT.Count, tr, esPublicorreo);

                                colsGT = new object[numDestinosPorTabla + 1];
                                colsGT[0] = descTramo;

                                for (int iDestino = 0; iDestino < numDestinosPorTabla; iDestino++)
                                {
                                    DestinoBE objDestino = destinos[(iTabla * numDestinosPorTablaInicial) + iDestino];

                                    var configGT = listaConfiguracionGrupoTramos.Where(t => t.CodDestino.Equals(objDestino.CodDestinoSAP) && t.Nombre.Equals(tramo.Nombre)).FirstOrDefault();
                                    Decimal valorDto = configGT != null ? configGT.DtoPC : 0;

                            if (objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoGrupoTramo))
                            {
                                        String dtoStr = decimal.Round(valorDto, redondeo).ToString() + valorCaracter;
                                        colsGT[iDestino + 1] = dtoStr;
                            }
                            else
                            {
                                        if (valorDto != 0)
                                            colsGT[iDestino + 1] = decimal.Round(valorDto, redondeo);
                                        
                            }

                                }

                            objExcel.AgregarFila(colsGT);
                        }

                            objExcel.TerminarDibujarTabla(false, EsTablaGrupoTramo: true);
                            primeraCeldaGrupoTramos = objExcel.Get_PrimeraCeldaTabla;
                            columnasRequeridasGrupoTramos = objExcel.Get_ColumnasRequeridas;
                            objExcel.DarFormatoTablaPrecioCierto(numDestinosPorTabla, numTramos: listaTramosGT.Count, esPublicorreo: esPublicorreo);
                    }
                    else
                    {
                        objExcel.BuscaYSustituye("$$TABLAGRTAMOS", string.Empty, true, false);
                    }

                    //Relleno la tabla como si hubiera destinos
                    object[] colsDestino;

                    //añadir primera fila con las cabeceras de los destinos                             
                        colsDestino = new object[numDestinosPorTabla + 1];
                        colsDescripcion = new object[numDestinosPorTabla + 1];
                    colsDestino[0] = string.Empty;
                        colsDescripcion[0] = string.Empty;
                    int destino = 1;
                        
                        //Es Informe de Precios DD
                        var esInformePreciosDD = (objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoDestino) && tarifasConDescuento);

                        if (!esInformePreciosDD)
                    {
                            colsDescripcion[0] = "Peso";
                        }

                        for (int iDestino = 0; iDestino < numDestinosPorTabla; iDestino++)
                        {
                            DestinoBE objDestino = listadoDestinos[(iTabla * numDestinosPorTablaInicial) + iDestino];

                            //Obtenemos la descripcion
                            var objDescDestino = listaDestinos.FirstOrDefault(d => d.CodDestinoSAP.Equals(objDestino.CodDestinoSAPSinZona));
                            var descDestino = (objDescDestino != null && !String.IsNullOrEmpty(objDescDestino.DescripcionDestino)) ? objDescDestino.DescripcionDestino : "";
                            descDestino = StringUtil.ToTitleCaseIfAllUpper(descDestino);

                            colsDestino[destino] = objDestino.CodDestinoSAPSinZona;
                            colsDescripcion[destino] = descDestino;
                        destino++;
                    }

                    int decimales = objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoGrupoTramo) ? 2 : 5;
                    objExcel.IniciarDibujarTabla(true, true);

                    objExcel.AgregarFila(colsDestino);
                        objExcel.AgregarFila(colsDescripcion, esAnchoFijo: true);

                    //variable que uso para meter los valores en la plantilla.tiene numDestinos+1 para meter en la primera casilla                            
                    int fila = 0;
                        Decimal pesoGrupoTramo = 0;

                        //foreach (TramoBE objTramo in auxTramo)
                        for (int t = 0; t < auxTramo.Count; t++)
                    {
                            colsDestino = new object[numDestinosPorTabla + 1];
                            
                            //Adaptamos la descripción al formato desde _ hasta _ 
                            TramoBE objTramo = auxTramo[t];

                            if (objTramo.CodTramo != null && objTramo.CodTramo.Contains("E"))
                                mostrarAclaracionExpediciones = true;

                            colsDestino[0] = GetDescripionTramo(objProductoBE, objTramo, listaTramosInfTarifas, auxTramo.Count, t, ref pesoGrupoTramo);

                            for (int columna = 0; columna < numDestinosPorTabla; columna++)
                        {
                                colsDestino[columna + 1] = matrizValores[fila, (iTabla * numDestinosPorTablaInicial) + columna];
                        }

                        objExcel.AgregarFila(colsDestino);

                        fila++;

                        if (fila >= numTramos)
                        {
                            break;
                        }
                    }

                        objExcel.TerminarDibujarTabla(false, false, (int?)maxDecimalesTarifa);
                        primeraCeldaTramos = objExcel.Get_PrimeraCeldaTabla;

                        objExcel.DarFormatoTablaPrecioCierto(numDestinosPorTabla, numTramos, esPublicorreo);

                        //Si no coinciden las columnas de descripción de tramos de las dos tablas, hacemos un shift hasta que coincidan
                        if (columnasRequeridasGrupoTramos != null && primeraCeldaGrupoTramos != null)
                    {
                            var maxColumnasRequeridasTablaGT = columnasRequeridasGrupoTramos[0];
                            var maxColumnasRequeridasTablaTramos = objExcel.Get_ColumnasRequeridas[0];

                            if (maxColumnasRequeridasTablaTramos > maxColumnasRequeridasTablaGT)
                            {
                                objExcel.AumentarTamanyoTablaGrupoTramos(primeraCeldaGrupoTramos,
                                                                         listaTramosGT.Count + 2,
                                                                         maxColumnasRequeridasTablaTramos - maxColumnasRequeridasTablaGT,
                                                                         maxColumnasRequeridasTablaGT);
                            }
                            else if (maxColumnasRequeridasTablaTramos < maxColumnasRequeridasTablaGT)
                            {
                                objExcel.AumentarTamanyoTablaGrupoTramos(primeraCeldaTramos,
                                                                         auxTramo.Count + 2,
                                                                         maxColumnasRequeridasTablaGT - maxColumnasRequeridasTablaTramos,
                                                                         maxColumnasRequeridasTablaTramos);
                            }
                        }



                    }
                    
                    //if (objProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Tramos)) && mostrarAclaracionExpediciones)
                    if (mostrarAclaracionExpediciones)
                    {                        
                        objExcel.EscribirTextoAlFinalDelDocumento(SimuladorResources.AclaracionKiloAdicional);
                    }

                    if (mostrarAclaracionExpediciones)
                    {
                        if (objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoDestino))
                        {
                            objExcel.EscribirTextoAlFinalDelDocumento(SimuladorResources.AclaracionKiloAdicional.Replace("60", "30"));
                        }
                        else
                        {
                            objExcel.EscribirTextoAlFinalDelDocumento(SimuladorResources.AclaracionKiloAdicional);
                        }
                    }


                    //Rellenamos la tabla de VA.
                    if (ListaTarifasVAReporte.Count > 0)
                    {
                        objExcel.IniciarDibujarTabla(true);
                        string[] colsVATitulo = new string[2];
                        colsVATitulo[0] = "LISTADO DE VALORES AÑADIDOS";
                        colsVATitulo[1] = string.Empty;
                        objExcel.AgregarFila(colsVATitulo, ajustarTextoColumna: false);
                        foreach (ReporteVABE objVA in ListaTarifasVAReporte)
                        {
                            string[] colsVA = new string[2];
                            colsVA[0] = objVA.Nombre;
                            colsVA[1] = objVA.Descripcion;
                            objExcel.AgregarFila(colsVA);
                        }
                        objExcel.TerminarDibujarTabla(false, false, null, true);
                    }

                    #endregion

                    #region Sustitución de tags

                    // Al contrario que en word, hacemos esto lo último para no tener que insertar la tabla entre medias

                    objExcel.InsertarDocumentoAlFinal(rutaLineaProducto);

                    //Sustituimos las etiquetas,
                    objEtiquetas.Add("$$TITULOREPORTE", textoParametro);
                    objEtiquetas.Add("$$CODIGOSAP", string.Format(CultureInfo.InvariantCulture, "{0}", objProductoBE.Descripcion));
                    objEtiquetas.Add("$$VALIDEZDESDE", fechaInicial);
                    objEtiquetas.Add("$$VALIDEZHASTA", fechaFinal);
                    objEtiquetas.Add("$$NOMBRECLIENTECOMERCIAL", nombreCliente);
                    objEtiquetas.Add("$$TABLAGRTAMOS", string.Empty);
                    objEtiquetas.Add("$$MAXPESOVOLUMETRICO", textoPesoVolumetrico);
                    objExcel.BuscaYSustituye(objEtiquetas, true, false);

                    #endregion

                    #region insertar pie de informe según linea de producto

                    InformacionDestinosBL objInfoDestinosBL = new InformacionDestinosBL();
                    Collection<InformacionDestinosBE> objListaDestinos = objInfoDestinosBL.ObtenerListadoInformacionDestinos(objProductoBE.CodProducto);
                    StringBuilder sb = new StringBuilder();
                    if (objListaDestinos != null)
                    {
                        int insertados = 1;
                        foreach (DestinoBE objDestino in listadoDestinos)
                        {
                            InformacionDestinosBE objDescripcion = objListaDestinos.FirstOrDefault(x => x.CodDestinoSAP.Equals(objDestino.CodDestinoSAP));

                            if ((objDescripcion != null) && (!string.IsNullOrWhiteSpace(objDescripcion.DescripcionDestino)))
                            {
                                if (insertados < numDestinos)
                                {
                                    sb.Append(string.Format(CultureInfo.InvariantCulture, "{0}: {1}, ", objDescripcion.CodDestinoSAP, objDescripcion.DescripcionDestino));
                                }
                                else
                                {
                                    sb.Append(string.Format(CultureInfo.InvariantCulture, "{0}: {1}. ", objDescripcion.CodDestinoSAP, objDescripcion.DescripcionDestino));
                                }
                                insertados++;
                            }
                        }

                        //En caso de que la lista de destinos del producto y de la DB difieran.
                        if(sb.ToString() != string.Empty) sb.Replace(',', '.', sb.Length - 2, 1);
                    }

                    //objExcel.BuscaYSustituye("$$OBSERVACIONESDESTINOS", sb.ToString(), true, false);
                    objExcel.BuscaYSustituye(objEtiquetas, true, false);

                    #endregion
                }

                #endregion

                #region Abrir el fichero

                if (objExcel.Abierto)
                {
                    objExcel.GuardarLibro();
                    objExcel.CerrarExcel();
                }

                if (!informeMultiple)
                {
                    ManagerExcel.AbrirExcelStandalone(ficheroTemporalExcel);
                }

                #endregion
            }
            catch (Exception ex)
            {
                RegistrarAccionesSimulador.GuardarExcepcion(ex);
            }
            finally
            {
                if (objExcel.Abierto)
                {
                    objExcel.CerrarExcel();
                }
            }
        }

        /// <summary>
        /// Genera el informe de tarifas/precios en formato PDF
        /// </summary>
        /// <param name="objProductoBE">productoBE del que se muestra el informe</param>
        /// <param name="objProductoOfertaBE">productoOfertaBE del que se muestra el informe</param>
        /// <param name="fechaInicial">fecha inicial de validez de la oferta</param>
        /// <param name="fechaFinal">fecha final de validez de la oferta</param>
        /// <param name="nombreCliente">Nombre del cliente al que pertenece la oferta</param>
        /// <param name="tarifasConDescuento">indica si queremos el informe de precios o el de tarifas</param>
        /// <param name="esUltimoInforme">indica si es el último producto de un informe múltimple</param>
        private void GenerarInformeModeloGrupoTramos(ProductoBE objProductoBE, ProductoOfertaBE objProductoOfertaBE, string fechaInicial, string fechaFinal, string nombreCliente, bool tarifasConDescuento, bool informeMultiple, string ficheroTemporalWord, string ficheroTemporalPdf, bool esUltimoInforme)
        {
            #region Variables

            //En caso de ser un producto con Destinos se rellena esta matriz de informacion
            string[,] matrizValores = null;

            //Listado de valores que se ingresan en la tabla de valores añadidos si corresponde
            Collection<ReporteVABE> ListaTarifasVAReporte = new Collection<ReporteVABE>();

            //Lista de etiquetas con su correspondiente valor que se sustituye en el documento word
            Dictionary<string, object> objEtiquetas = new Dictionary<string, object>();

            //Contiene la ruta de la plantilla que se usa para generar el reporte de tarifas
            string rutaPlantilla = string.Format(Utils.GetValorFromAppConfig(AppSettingsEnum.RutaPlantillaInformeTarifas), AppDomain.CurrentDomain.BaseDirectory);

            string rutaLineaProducto = string.Format(Utils.GetValorFromAppConfig(AppSettingsEnum.RutaPlantillaLineaProducto), AppDomain.CurrentDomain.BaseDirectory, objProductoBE.PlantillaInformeTarifasPrecios);
            string rutaPlantillaVA = string.Format(Utils.GetValorFromAppConfig(AppSettingsEnum.RutaPlantillaInformeTarifasVA), AppDomain.CurrentDomain.BaseDirectory);
            string rutaPlantillaGT = string.Format(Utils.GetValorFromAppConfig(AppSettingsEnum.RutaPlantillaInformeTarifasGT), AppDomain.CurrentDomain.BaseDirectory);
            string descripciones;

            //Se activa para los casos que se quiere mostrar la aclaración *Para expediciones.
            Boolean mostrarAclaracionExpediciones = false;

            // Guardamos si es precio cierto o tipo descuento
            bool esTipoPrecioCierto = (objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoPrecioCierto) ||
                                       objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorPrecioCiertoGrupoTramo));
            
            // Si no es un informe múltiple o no tiene nombre, poner nombre por defecto
            if (ficheroTemporalWord.Equals(string.Empty) || !informeMultiple)
            {
                ficheroTemporalWord = Path.Combine(System.IO.Path.GetTempPath(), "tarifas.docx");
            }

            String ficheroDestino = string.Empty;

            //instancia del objeto word con el que vamos  trabajar para generar el informe de tarifas
            ManagerWord objWord = new ManagerWord(ficheroTemporalWord, false);

            //Lista de tramos que usamos para, al generar el informe, mostrar la cabecera de los tramos.
            List<TramoBE> auxTramo = new List<TramoBE>();

            Collection<TramoBE> listaTramosEliminar = new Collection<TramoBE>();

            #endregion

            try
            {
                #region Obtencion Datos

                //Obtenemos los datos que ha ingresado el usuario.
                ConfiguracionGruposTramoBL configuracionGruposTramoBL = new ConfiguracionGruposTramoBL();
                ConfiguracionProductosBL configProducto = new ConfiguracionProductosBL();

                Collection<GrupoTramoBE> listaConfiguracionGrupoTramos = configuracionGruposTramoBL.ObtenerListaGruposTramoOferta(objProductoBE, objProductoOfertaBE.idProductoOferta);
                Collection<ConfiguracionTramoOfertaBE> listaConfiguracionTramo = configProducto.ObtenerConfiguracionTramoOferta(objProductoOfertaBE.idProductoOferta);
                Collection<InformacionDestinosBE> listaDestinos = new InformacionDestinosBL().ObtenerListadoInformacionDestinos(objProductoBE.CodProducto);
                List<TramoInformeBE> listaTramosInfTarifas = new TramoInformeBL().ObtenerTramosInformeProducto(objProductoBE.CodProducto);
                Collection<ConfiguracionDestinoOfertaBE> listaConfiguracionDestino = configProducto.ObtenerConfiguracionDestinoOferta(objProductoOfertaBE.idProductoOferta);

                //Se obtienen los destinos y tramos para los descuentos sobre las tarifas
                List<String> listaDestinosGT = listaConfiguracionGrupoTramos.OrderBy(t => t.OrdenDestino).GroupBy(t => t.CodDestino).Select(t => t.First()).Select(t => t.CodDestino).ToList();
                List<GrupoTramoBE> listaTramosGT = listaConfiguracionGrupoTramos.OrderBy(t => t.OrdenDestino).OrderBy(t => t.CodTramoInicialDecimal).GroupBy(t => t.Nombre).Select(t => t.First()).ToList();

                //Vemos si es de publicorreo
                Boolean esPublicorreo = objProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Publicorreo));
                //Producto con destinos
                int numTramos = 0;


                var listadoDestinos = objProductoBE.Destinos.Where(x => x.Tramos.Count > 0).OrderBy(t => t.Orden).ToList();

                if (objProductoBE.Internacional)
                {
                    listadoDestinos = (from d in listadoDestinos
                                       join c in listaConfiguracionDestino on d.idDestino equals c.idDestino
                                       where c.Distribucion.HasValue && c.Distribucion.Value > 0
                                       select d).ToList();

                }

                int numDestinos = listadoDestinos.Count();

                //obtenemos el número máximo de tramos de entre los destinos. Ademas nos guardamos los tramos para la cabecera de fila de la tabla del informe. 
                foreach (DestinoBE destino in listadoDestinos)
                {
                    if (numTramos < destino.Tramos.Count)
                    {
                        numTramos = destino.Tramos.Count;
                        //auxTramo = destino.Tramos;
                        TramoBE[] copiaTramos = new TramoBE[numTramos];
                        destino.Tramos.CopyTo(copiaTramos, 0);
                        auxTramo = copiaTramos.ToList();
                    }
                }

                //En este punto ya tenemos en auxTramo la colleccion de tramos que hay que mostrar.
                foreach (TramoBE item in auxTramo)
                {
                    //Eliminamos los tramos que no estén en ningun grupo de tramo
                    bool estaTramoEnGrupos = false;
                    foreach (GrupoTramoBE itemGrTamo in listaConfiguracionGrupoTramos)
                    {
                        if ((item.CodTramoDecimal >= itemGrTamo.CodTramoInicialDecimal) && (item.CodTramoDecimal <= itemGrTamo.CodTramoFinalDecimal))
                        {
                            estaTramoEnGrupos = true;
                            break;
                        }
                    }
                    if (!estaTramoEnGrupos)
                    {
                        listaTramosEliminar.Add(item);
                    }
                }

                if (auxTramo.Count == listaTramosEliminar.Count)
                {
                    listaTramosEliminar.Clear();
                }

                numTramos = auxTramo.Count;
                Collection<TramoBE> listaTramosInforme = auxTramo.ToList().ToCollection();

                //Eliminamos los tramos que no deben insertarse
                foreach (TramoBE item in listaTramosEliminar)
                {
                    listaTramosInforme.Remove(item);
                    numTramos--;
                }

                matrizValores = new string[numTramos, numDestinos];
                int i = 0;
                int j = 0;

                //Para los productos S0134 y S0235 y los destinos Z7, Z8 y Z9 No se deben mostrar los tramos de expediciones
                //JCNS. MOSTRAR TRAMO. LO CAMBIO DE SITIO
                //bool DebeMostrarTramo = true;

                int maxDecimalesTarifa = 2;

                foreach (DestinoBE destino in listadoDestinos)
                {
                    j = 0;
                    foreach (TramoBE tramo in destino.Tramos)
                    {
                        //Para los productos S0134, S0132 y S0235 y los destinos Z7, Z8 y Z9 No se deben mostrar los tramos de expediciones
                        //JCNS. MOSTRAR TRAMO. LO CAMBIO DE SITIO
                        bool DebeMostrarTramo = true;

                        if ((objProductoBE.CodProducto.Equals("S0236") || objProductoBE.CodProducto.Equals("S0133") || objProductoBE.CodProducto.Equals("S0134") || objProductoBE.CodProducto.Equals("S0235") || objProductoBE.CodProducto.Equals("S0132")) && (destino.CodDestinoSAP.Equals("Z7") || destino.CodDestinoSAP.Equals("Z8") || destino.CodDestinoSAP.Equals("Z9")) && (tramo.CodTramo.StartsWith("E")))
                        {
                            DebeMostrarTramo = false;
                        }

                        TramoBE auxEliminar = listaTramosEliminar.FirstOrDefault(x => x.CodTramo.Equals(tramo.CodTramo));
                        if (auxEliminar == null)
                        {
                            TramoBE aux = auxTramo.FirstOrDefault(x => x.CodTramo.Equals(tramo.CodTramo));
                            if (aux != null)
                            {
                                string tarifaTramo = Math.Round(tramo.Tarifa, 5).ToString();

                                if (tarifasConDescuento)
                                {
                                    double auxTarifa = 0;


                                    //MMunoz FIX: No cogía el descuento del grupo, si no del tramo concreto. Lo hemos modificado, para que saque
                                    //el valor de Dto o PC del grupo.

                                    //Se obtiene la configuracion del grupo de tramos                                    
                                    GrupoTramoBE configGrupoTramo = listaConfiguracionGrupoTramos.FirstOrDefault(x => x.CodDestino.Equals(destino.CodDestinoSAP)
                                                        && x.CodTramoInicialDecimal <= tramo.CodTramoDecimal
                                                        && x.CodTramoFinalDecimal >= tramo.CodTramoDecimal
                                                        && x.idProductoOferta.Equals(objProductoOfertaBE.idProductoOferta));

                                    //Se obtiene la configuración del tramo
                                    ConfiguracionTramoOfertaBE configTramo = listaConfiguracionTramo.FirstOrDefault(x => x.idTramo.Equals(tramo.idTramo));

                                    if (configTramo != null)
                                    {
                                        //FIX: intentamos sacar el Dto o el PC del grupo. Si no hay, lo obtenemos de los valores del tramo.                               
                                        var descuentoFinal = (configGrupoTramo != null) ? configGrupoTramo.DtoPC : configTramo.DescuentoFinal.Value;
                                        var precioCierto = (configGrupoTramo != null) ? configGrupoTramo.DtoPC : configTramo.PrecioCierto.Value;

                                        if ((objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoGrupoTramo)) &&
                                                configTramo.DescuentoFinal.HasValue && double.TryParse(descuentoFinal.ToString(), out auxTarifa))
                                        {
                                            auxTarifa = tramo.Tarifa - (tramo.Tarifa * auxTarifa / 100);
                                            tarifaTramo = Math.Round(auxTarifa, 5).ToString();
                                        }
                                        else if ((objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorPrecioCiertoGrupoTramo)) &&
                                            configTramo.PrecioCierto.HasValue && double.TryParse(precioCierto.ToString(), out auxTarifa))
                                        {
                                            if (!auxTarifa.Equals(0))
                                            {
                                                tarifaTramo = Math.Round(auxTarifa, 5).ToString();
                                            }
                                        }
                                    }
                                }

                                if (DebeMostrarTramo)
                                {
                                    decimal argument = decimal.Parse(tarifaTramo);
                                    int count = BitConverter.GetBytes(decimal.GetBits(argument)[3])[2];

                                    if (count >= maxDecimalesTarifa)
                                        maxDecimalesTarifa = count;

                                    matrizValores[j, i] = tarifaTramo;
                                }
                                else
                                {
                                    matrizValores[j, i] = string.Empty;
                                }

                                j++;
                                DebeMostrarTramo = true;
                            }
                        }
                    }
                    i++;
                }

                //Asignamos el decimal más grande de los posibles
                for (i = 0; i < matrizValores.GetLength(0); i++)
                {
                    for (j = 0; j < matrizValores.GetLength(1); j++)
                    {
                        if (!String.IsNullOrEmpty(matrizValores[i, j]))
                        {
                            double tarifa = double.Parse(matrizValores[i, j]);
                            matrizValores[i, j] = tarifa.ToString("N" + maxDecimalesTarifa) + '€';
                        }
                    }
                }

                //Se buscan los datos de los VA            
                ListaTarifasVAReporte = this.CrearTablaVA(objProductoBE.idProducto, objProductoOfertaBE.idProductoOferta, tarifasConDescuento);

                //Se busca el valor máximo de peso volumétrico
                int maxPesoVolumetrico; //40  o 60

                //Vemos si se ha definido un peso volumétrico máximo especifico para el tipo de producto actual (CodAnexoSAP)
                maxPesoVolumetrico = new ProductoSAPBL().ObtenerPesoVolumetricoMaxProducto(objProductoBE.CodProducto);

                //Si no hay valor máximo registrado, se coge el valor por defecto (40 Kg).
                if (maxPesoVolumetrico == 0) maxPesoVolumetrico = 40;
                String textoPesoVolumetrico = string.Format(SimuladorResources.PesoVolumetricoMaximoInforme, maxPesoVolumetrico);

                #endregion

                #region Generacion Plantilla

                //Copiamos la plantilla en el fichero temporal.
                File.Copy(rutaPlantilla, ficheroTemporalWord, true);

                int numTablasCrear = 1;
                int numDestinosPorTabla = 15;

                //Paqueteria kg, si no en g. Miramos si su modelo de descuento es paquetería, tramos, o aparece como paquete en el doc. de tarifas                        
                //Si es Publicorreo óptimo, no lo consideramos como paquetería     
                bool paqueteria = !ModeloDescuentoEnum.GetPaqueteriasQueSonPublicorreos().Contains(objProductoBE.CodProducto) && objProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Paqueteria)) ||
                                  objProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Tramos)) ||
                                  listaTramosInfTarifas.Count > 0;

                //Si es de paquetería internacional
                if (objProductoBE.Internacional)
                {
                    numTablasCrear = (int)Math.Ceiling(numDestinos / 15.0);

                    if (numTablasCrear != 1)
                    {
                        numDestinosPorTabla = 15;
                    }
                    else
                    {
                        numDestinosPorTabla = numDestinos;
                    }

                }
                //Si es normal
                else
                {
                    numTablasCrear = 1;
                    numDestinosPorTabla = numDestinos;
                }


                //Abrimos el fichero
                objWord.AbrirFichero();
                if (objWord.Abierto)
                {
                    objWord.ActivarModoWord();

                    #region Etiquetas word

                    if (numDestinos <= 9)
                    {
                        objWord.CambiarOrientacionVertical();
                    }
                    else if (numDestinos > 16)
                    {
                        objWord.AgrandarAnchuraWord(numDestinos - 16);
                    }

                    string textoParametro = string.Empty;
                    if (!tarifasConDescuento)
                    {
                        textoParametro = SimuladorResources.TituloReportTarifas;
                    }
                    else
                    {
                        textoParametro = SimuladorResources.TituloReportPrecios;
                    }

                    //Sustituimos las etiquetas,
                    //String saltoLinea = "\f"; //(ListaTarifasVAReporte.Count > 0) ? "\f" : String.Empty;

                    //if (numTablasCrear > 1 || (listaTramosInfTarifas.Count > 0 && ListaTarifasVAReporte.Count > 0) || (paqueteria && (!objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoDestino) || listaTramosInfTarifas.Count == 0))) //Paqueteria
                    //{
                    //    saltoLinea = "\f";
                    //}
                    //else
                    //{
                    //    saltoLinea = "";
                    //}
                    //saltoLinea = "";

                    objEtiquetas.Add("$$TITULOREPORTE", textoParametro);
                    objEtiquetas.Add("$$CODIGOSAP", string.Format(CultureInfo.InvariantCulture, "{0}", objProductoBE.Descripcion));
                    objEtiquetas.Add("$$VALIDEZDESDE", fechaInicial);
                    objEtiquetas.Add("$$VALIDEZHASTA", fechaFinal);
                    objEtiquetas.Add("$$MAXPESOVOLUMETRICO", textoPesoVolumetrico);
                    objEtiquetas.Add("$$NOMBRECLIENTECOMERCIAL", nombreCliente);
                    #endregion

                    #region Rellenar las tablas

                    //INICIAR CAMBIOS

                    int indiceTabla = 1;
                    int numDestinosPorTablaInicial = numDestinosPorTabla;

                    //Duplicamos la primera tabla tantas veces como tablas a crear                    
                    for (int z = 0; z < (2 * numTablasCrear) - 2; z++)
                    {
                        objWord.CopiarTabla(1); //Tramos
                    }

                    //Creamos T tablas, de acuerdo al tipo de producto
                    for (int iTabla = 0; iTabla < numTablasCrear; iTabla++)
                    {
                        if (listaConfiguracionGrupoTramos.Count > 0 && tarifasConDescuento)
                        {
                            //Si es la última tabla, nos aseguramos del nº de columnas a mostrar    
                            if ((iTabla == numTablasCrear - 1) && (numDestinos % 15 != 0))
                            {
                                numDestinosPorTabla = numDestinos % 15;
                            }

                            //Rellenamos la tabla
                            objWord.SeleccionarTabla((iTabla * 2) + 1);
                            objWord.InsertarTablaParteSuperior(indiceTabla);

                            string textoRemplazo = "PRECIO CIERTO";
                            string valorCaracter = " €";
                            int redondeo = 5;

                            if (objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoGrupoTramo))
                            {
                                textoRemplazo = "DESCUENTO";
                                valorCaracter = " %";
                                redondeo = 2;
                            }

                            string tablaDestinosGT = "\t";
                            descripciones = esTipoPrecioCierto ? "GRUPOS DE TRAMOS\t" : "GRUPOS DE TRAMOS$SALTODescuento aplicado sobre tarifa\t";

                            Action<String> anyadirValTablaGT = new Action<String>(t => tablaDestinosGT += t + '\t');

                            for (int iDestino = 0; iDestino < numDestinosPorTabla; iDestino++)
                            {
                                DestinoBE objDestino = listadoDestinos[(iTabla * numDestinosPorTablaInicial) + iDestino];

                                //Obtenemos la descripcion
                                var objDescDestino = listaDestinos.FirstOrDefault(d => d.CodDestinoSAP.Equals(objDestino.CodDestinoSAPSinZona));
                                var descDestino = (objDescDestino != null && !String.IsNullOrEmpty(objDescDestino.DescripcionDestino)) ? objDescDestino.DescripcionDestino : "";
                                descDestino = StringUtil.ToTitleCaseIfAllUpper(descDestino);
                                // Añadimos cada destino
                                anyadirValTablaGT(objDestino.CodDestinoSAPSinZona);
                                descripciones += descDestino + '\t';
                            }

                            // Borramos el caracter extra insertado en el for
                            tablaDestinosGT = tablaDestinosGT.Substring(0, tablaDestinosGT.Length - 1);
                            descripciones = descripciones.Substring(0, descripciones.Length - 1);
                            tablaDestinosGT += '\n';
                            tablaDestinosGT += descripciones + '\n';

                            //Rellenar los datos de la tabla                                                                      
                            for (int tr = 0; tr < listaTramosGT.Count; tr++)
                            {
                                //var tramoFormateado = tramo.Replace("de", "Más de").Replace(" a ", " hasta ") + ".";
                                GrupoTramoBE tramo = listaTramosGT[tr];
                                String descTramo = GetDescripionGrupoTramo(objProductoBE, tramo, listaTramosInfTarifas, listaTramosGT.Count, tr, esPublicorreo);

                                anyadirValTablaGT(descTramo);

                                //foreach (DestinoBE objDestino in objProductoBE.Destinos.Where(x => x.Tramos.Count > 0).OrderBy(x => x.Orden).ToList())
                                //{
                                for (int iDestino = 0; iDestino < numDestinosPorTabla; iDestino++)
                                {
                                    DestinoBE objDestino = listadoDestinos[(iTabla * numDestinosPorTablaInicial) + iDestino];
                                    var configGT = listaConfiguracionGrupoTramos.Where(t => t.CodDestino.Equals(objDestino.CodDestinoSAP) && t.Nombre.Equals(tramo.Nombre)).FirstOrDefault();
                                    Decimal valorDto = configGT != null ? configGT.DtoPC : 0;
                                    String dtoStr = decimal.Round(valorDto, redondeo).ToString("N2") + valorCaracter;

                                    if (esTipoPrecioCierto && valorDto == 0)
                                    {
                                        dtoStr = String.Empty;
                                    }

                                    anyadirValTablaGT(dtoStr);
                                }

                                // Borramos el caracter extra insertado en el for
                                tablaDestinosGT = tablaDestinosGT.Substring(0, tablaDestinosGT.Length - 1);
                                tablaDestinosGT += '\n';
                            }

                            // Borramos el caracter extra insertado en el for
                            tablaDestinosGT = tablaDestinosGT.Substring(0, tablaDestinosGT.Length - 1);

                            objWord.EscribirTablaDeString(tablaDestinosGT, true, null);
                            objWord.PonerTablaCabeceraRepetida();
                            objWord.DarFormatoTablaGrupoDestino();

                            //sustituir cabecera columna
                            objWord.BuscaYSustituye("$$TIPODESCUENTO", textoRemplazo, true, false);
                            objWord.BuscaYSustituye("$$TABLAGRTAMOS", string.Empty, true, false);
                            //una vez rellenada la tabla                         
                            indiceTabla++;

                            //Seleccionamos la tabla de tarifas.
                            objWord.SeleccionarTabla((iTabla * 2) + 2);
                        }
                        else
                        {
                            objWord.BuscaYSustituye("$$TABLAGRTAMOS", string.Empty, true, false);

                            //Seleccionamos la tabla de tarifas.
                            objWord.SeleccionarTabla((iTabla * 2) + 1);
                        }

                        string tablaEnString = "\t"; // La primera celda de la cabecera es vacía
                        descripciones = "\t";

                        //Es Informe de Precios DD
                        var esInformePreciosDD = (objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoDestino) && tarifasConDescuento);

                        if (!esInformePreciosDD)
                        {
                            descripciones = "Peso\t";
                        }

                        //foreach (DestinoBE objDestino in objProductoBE.Destinos.Where(x => x.Tramos.Count > 0).OrderBy(x => x.Orden).ToList())
                        //{
                        for (int iDestino = 0; iDestino < numDestinosPorTabla; iDestino++)
                        {
                            DestinoBE objDestino = listadoDestinos[(iTabla * numDestinosPorTablaInicial) + iDestino];
                            //Obtenemos la descripcion
                            var objDescDestino = listaDestinos.FirstOrDefault(d => d.CodDestinoSAP.Equals(objDestino.CodDestinoSAPSinZona));
                            var descDestino = (objDescDestino != null && !String.IsNullOrEmpty(objDescDestino.DescripcionDestino)) ? objDescDestino.DescripcionDestino : "";
                            descDestino = StringUtil.ToTitleCaseIfAllUpper(descDestino);

                            // Añadimos cada destino
                            tablaEnString += objDestino.CodDestinoSAPSinZona + '\t';
                            descripciones += descDestino + '\t';
                        }
                        // Borramos el caracter extra insertado en el for
                        tablaEnString = tablaEnString.Substring(0, tablaEnString.Length - 1);
                        tablaEnString += '\n';
                        tablaEnString += descripciones.Substring(0, descripciones.Length - 1);
                        tablaEnString += '\n';

                        int fila = 0;
                        Decimal pesoGrupoTramo = 0;

                        //foreach (TramoBE objTramo in listaTramosInforme)
                        //{
                        for (int t = 0; t < listaTramosInforme.Count; t++)
                        {
                            string celda = "";

                            //// Primera celda de cada tramo, contiene la descripción del tramo
                            //tablaEnString += (objTramo.Descripcion.Replace(" GRS", "")).Replace(".", "") + '\t';

                            //Adaptamos la descripción al formato desde _ hasta _ 
                            TramoBE objTramo = listaTramosInforme[t];
                            string descTramo = GetDescripionTramo(objProductoBE, objTramo, listaTramosInfTarifas, listaTramosInforme.Count, t, ref pesoGrupoTramo);

                            if (objTramo != null && objTramo.CodTramo.Contains("E"))
                                mostrarAclaracionExpediciones = true;

                            tablaEnString += descTramo;

                            //for (int columna = 0; columna < numDestinos; columna++)
                            //{
                            for (int columna = 0; columna < numDestinosPorTabla; columna++)
                            {
                                // Valores en € del tramo
                                celda = matrizValores[fila, (iTabla * numDestinosPorTablaInicial) + columna];
                                tablaEnString += celda + '\t';
                            }

                            // Borramos el caracter extra insertado en el for
                            tablaEnString = tablaEnString.Substring(0, tablaEnString.Length - 1);
                            tablaEnString += '\n';

                            fila++;

                            if (fila >= numTramos)
                            {
                                break;
                            }
                        }

                        // Borramos el caracter extra insertado en el for
                        tablaEnString = tablaEnString.Substring(0, tablaEnString.Length - 1);

                        objWord.EscribirTablaDeString(tablaEnString, true, maxDecimalesTarifa);

                        objWord.PonerTablaCabeceraRepetida();
                        objWord.DarFormatoTablaPrecioCierto();

                    }

                    // if (objProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Tramos)) && mostrarAclaracionExpediciones)
                    if (mostrarAclaracionExpediciones)
                    {
                        objWord.EscribirTextoAlFinalDelDocumento(SimuladorResources.AclaracionKiloAdicional);
                    }

                    //if (auxTramo.Count > 25)
                    //{
                    //    objWord.InsertarSaltoPaginaAlFinalDocumentoSiNecesario();
                    //}

                    objWord.EscribirTextoAlFinalDelDocumento("$$TABLAVA");
                    indiceTabla++;

                    //Rellenamos la tabla de VA.
                    if (ListaTarifasVAReporte.Count > 0)
                    {
                        //insertamos la plantilla de la tabla de VA
                        objWord.InsertarDocumentoSustituyendoTexto(rutaPlantillaVA, "$$TABLAVA", (int)FormasSustitucionTexto.CopiarPegarConFormato);

                        objWord.SeleccionarTabla(indiceTabla);
                        string[] colsVA = new string[2];
                        foreach (ReporteVABE objVA in ListaTarifasVAReporte)
                        {
                            colsVA[0] = objVA.Nombre;
                            colsVA[1] = objVA.Descripcion;
                            objWord.AgregarFilaSinFormatoATabla(colsVA);
                        }
                        objWord.PonerTablaCabeceraRepetida();
                        objWord.DarFormatoTablaVA();
                    }
                    else
                    {
                        objWord.BuscaYSustituye("$$TABLAVA", "$$OBSERVACIONES", true, false);
                    }


                    //FIN CAMBIOS

                    #endregion

                    #region insertar pie de informe según linea de producto

                    int numPaginas = objWord.GetPagesNumber();
                    objWord.InsertarDocumentoSustituyendoTexto(rutaLineaProducto, "$$OBSERVACIONES", (int)FormasSustitucionTexto.CopiarPegarConFormato);
                    int numPaginasPost = objWord.GetPagesNumber();
                    String saltoLinea = (numPaginasPost > numPaginas) ? "\f" : "";
                    objWord.BuscaYSustituye("$$SALTOLINEA", saltoLinea, true, false);   

                    InformacionDestinosBL objInfoDestinosBL = new InformacionDestinosBL();
                    Collection<InformacionDestinosBE> objListaDestinos = objInfoDestinosBL.ObtenerListadoInformacionDestinos(objProductoBE.CodProducto);
                    StringBuilder sb = new StringBuilder();
                    if (objListaDestinos != null)
                    {
                        int insertados = 1;
                        foreach (DestinoBE objDestino in listadoDestinos)
                        {
                            InformacionDestinosBE objDescripcion = objListaDestinos.FirstOrDefault(x => x.CodDestinoSAP.Equals(objDestino.CodDestinoSAP));

                            if ((objDescripcion != null) && (!string.IsNullOrWhiteSpace(objDescripcion.DescripcionDestino)))
                            {
                                if (insertados < numDestinos)
                                {
                                    sb.Append(string.Format(CultureInfo.InvariantCulture, "{0}: {1}, ", objDescripcion.CodDestinoSAP, objDescripcion.DescripcionDestino));
                                }
                                else
                                {
                                    sb.Append(string.Format(CultureInfo.InvariantCulture, "{0}: {1}. ", objDescripcion.CodDestinoSAP, objDescripcion.DescripcionDestino));
                                }
                                insertados++;
                            }
                        }

                        //En caso de que la lista de destinos del producto y de la DB difieran.
                        if (sb.ToString() != string.Empty) sb.Replace(',', '.', sb.Length - 2, 1);

                    }
                    //objWord.BuscaYSustituye("$$OBSERVACIONESDESTINOS", sb.ToString(), true, false);
                    objWord.BuscaYSustituye(objEtiquetas, true, false);


                    if (informeMultiple && !esUltimoInforme)
                        //objWord.InsertarSaltoPaginaAlFinalDocumento();
                        objWord.InsertarCambioSeccionAlFinalDocumento();

                    #endregion

                    #region Generar Word

                    //Ya tenemos generado el documento, ahora lo guardamos en Word.                    
                    ficheroDestino = string.Empty;

                    if (tarifasConDescuento)
                    {
                        ficheroDestino = Path.Combine(System.IO.Path.GetTempPath(), string.Format(CultureInfo.InvariantCulture, "{0}_{7}_{1}{2}{3}{4}{5}{6}.docx", "InformePrecios", DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(),
                            DateTime.Now.Day.ToString(), DateTime.Now.Hour.ToString(), DateTime.Now.Minute.ToString(), DateTime.Now.Second.ToString(), objProductoBE.CodProducto));
                    }
                    else
                    {
                        //ficheroDestino = Path.Combine(System.IO.Path.GetTempPath(), string.Format(CultureInfo.InvariantCulture, "{0}_{1}.docx", "InformeTarifas", objProductoBE.CodProducto));
                        ficheroDestino = Path.Combine(System.IO.Path.GetTempPath(), string.Format(CultureInfo.InvariantCulture, "{0}_{7}_{1}{2}{3}{4}{5}{6}.docx", "InformeTarifas", DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(),
                            DateTime.Now.Day.ToString(), DateTime.Now.Hour.ToString(), DateTime.Now.Minute.ToString(), DateTime.Now.Second.ToString(), objProductoBE.CodProducto));
                                                
                        if (File.Exists(ficheroDestino))
                        {
                            //si existe puede que esté abierto
                            //cerrar el proceso
                            //Process objPDF = Process.GetProcessesByName("winword").FirstOrDefault(x => x.MainWindowTitle.Contains(string.Format(CultureInfo.InvariantCulture, "{0}_{1}.docx", "InformeTarifas", objProductoBE.CodProducto)));
                            //if (objPDF != null)
                            //{
                            //    objPDF.Kill();
                            //}
                            }
                        }
                    //Guardamos el fichero
                    if (informeMultiple)
                    {
                        objWord.GuardarComo(ficheroTemporalPdf);
                    }
                    else
                    {
                        objWord.GuardarComo(ficheroDestino);
                    }

                    #endregion
                }

                #endregion

            }
            catch (Exception ex)
            {
                RegistrarAccionesSimulador.GuardarExcepcion(ex);
            }
            finally
            {
                if (objWord.Abierto)
                {
                    objWord.CerrarWord();
                }

                if (!informeMultiple)
                {
                    //Una vez creado el fichero en temporal se abre.
                    if (System.IO.File.Exists(ficheroDestino))
                    {
                        System.Diagnostics.Process proc = new System.Diagnostics.Process();
                        proc.EnableRaisingEvents = false;
                        proc.StartInfo.FileName = ficheroDestino;
                        proc.Start();
            }
        }
            }
        }

        /// <summary>
        /// Genera el informe de tarifas/precios en formato PDF
        /// </summary>
        /// <param name="objProductoBE">productoBE del que se muestra el informe</param>
        /// <param name="objProductoOfertaBE">productoOfertaBE del que se muestra el informe</param>
        /// <param name="fechaInicial">fecha inicial de validez de la oferta</param>
        /// <param name="fechaFinal">fecha final de validez de la oferta</param>
        /// <param name="nombreCliente">Nombre del cliente al que pertenece la oferta</param>
        /// <param name="tarifasConDescuento">indica si queremos el informe de precios o el de tarifas</param>
        /// <param name="informeMultiple">Indica si se debe crear un informe orientado a la generación de uno múltiple (sin guardar PDF y guardando Words)</param>
        /// <param name="esUltimoInforme">indica si es el último producto de un informe múltimple</param>
        private void GenerarInformeEstandar(ProductoBE objProductoBE, ProductoOfertaBE objProductoOfertaBE, string fechaInicial, string fechaFinal, string nombreCliente, bool tarifasConDescuento, bool informeMultiple, string ficheroTemporalWord, string ficheroTemporalPdf, bool esUltimoInforme)
        {
            #region Variables

            //En caso de ser un producto con Destinos se rellena esta matriz de informacion
            string[,] matrizValores = null;

            //Listado de valores que se ingresan en la tabla de valores añadidos si corresponde
            Collection<ReporteVABE> ListaTarifasVAReporte = new Collection<ReporteVABE>();

            //Lista de etiquetas con su correspondiente valor que se sustituye en el documento word
            Dictionary<string, object> objEtiquetas = new Dictionary<string, object>();

            //Contiene la ruta de la plantilla que se usa para generar el reporte de tarifas
            string rutaPlantilla = string.Format(Utils.GetValorFromAppConfig(AppSettingsEnum.RutaPlantillaInformeTarifas), AppDomain.CurrentDomain.BaseDirectory);

            string rutaLineaProducto = string.Format(Utils.GetValorFromAppConfig(AppSettingsEnum.RutaPlantillaLineaProducto), AppDomain.CurrentDomain.BaseDirectory, objProductoBE.PlantillaInformeTarifasPrecios);
            string rutaPlantillaVA = string.Format(Utils.GetValorFromAppConfig(AppSettingsEnum.RutaPlantillaInformeTarifasVA), AppDomain.CurrentDomain.BaseDirectory);
            string ficheroDestino = string.Empty;

            //string rutaPlantillaDimExtra = string.Format(Utils.GetValorFromAppConfig(AppSettingsEnum.RutaDimensionesExtra), AppDomain.CurrentDomain.BaseDirectory);

            //Se activa para los casos que se quiere mostrar la aclaración *Para expediciones.
            Boolean mostrarAclaracionExpediciones = false;

            // Si no es un informe múltiple o no tiene nombre, poner nombre por defecto
            if (ficheroTemporalWord.Equals(string.Empty) || !informeMultiple)
            {
                ficheroTemporalWord = Path.Combine(System.IO.Path.GetTempPath(), "tarifas.docx");
            }

            //instancia del objeto word con el que vamos  trabajar para generar el informe de tarifas



            ManagerWord objWord = new ManagerWord(ficheroTemporalWord, false);

            //Lista de tramos que usamos para, al generar el informe, mostrar la cabecera de los tramos.
            List<TramoBE> auxTramo = new List<TramoBE>();

            Collection<TramoBE> listaTramosEliminar = new Collection<TramoBE>();

            #endregion

            try
            {
                #region Obtencion Datos

                //Obtenemos los datos que ha ingresado el usuario.
                ConfiguracionProductosBL configProducto = new ConfiguracionProductosBL();
                Collection<ConfiguracionDestinoOfertaBE> listaConfiguracionDestino = configProducto.ObtenerConfiguracionDestinoOferta(objProductoOfertaBE.idProductoOferta);
                Collection<ConfiguracionTramoOfertaBE> listaConfiguracionTramo = configProducto.ObtenerConfiguracionTramoOferta(objProductoOfertaBE.idProductoOferta);
                Collection<InformacionDestinosBE> listaDestinos = new InformacionDestinosBL().ObtenerListadoInformacionDestinos(objProductoBE.CodProducto);
                List<TramoInformeBE> listaTramosInfTarifas = new TramoInformeBL().ObtenerTramosInformeProducto(objProductoBE.CodProducto);

                //Para no estar realizando la consulta del numero de tramos guardamos el valor la primera vez.
                var listadoDestinos = objProductoBE.Destinos.Where(x => x.Tramos.Count > 0).OrderBy(t => t.Orden).ToList();

                if (objProductoBE.Internacional)
                {
                    listadoDestinos = (from d in listadoDestinos
                                       join c in listaConfiguracionDestino on d.idDestino equals c.idDestino
                                       where c.Distribucion.HasValue && c.Distribucion.Value > 0
                                       select d).ToList();

                }

                int numDestinos = listadoDestinos.Count();

                //Producto con destinos
                int numTramos = 0;
                //obtenemos el número máximo de tramos de entre los destinos. Ademas nos guardamos los tramos para la cabecera de fila de la tabla del informe. 
                foreach (DestinoBE destino in listadoDestinos)
                {
                    if (numTramos < destino.Tramos.Count)
                    {
                        numTramos = destino.Tramos.Count;
                        //auxTramo = destino.Tramos;
                        TramoBE[] copiaTramos = new TramoBE[numTramos];
                        destino.Tramos.CopyTo(copiaTramos, 0);
                        auxTramo = copiaTramos.ToList();
                    }
                }

                decimal limiteTramos;

                if (objProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Publicorreo)))
                {
                    limiteTramos = 250;
                    foreach (TramoBE item in auxTramo.OrderBy(x => x.CodTramo))
                    {
                        if (item.CodTramoDecimal > limiteTramos)
                        {
                            //[FIX][MMUNOZ] Quieren que se muestren todos los valores de los tramos, tengan distribución o no
                            Collection<ConfiguracionTramoOfertaBE> auxRellenos = listaConfiguracionTramo.Where(x => (x.CodTramo.Equals(item.CodTramo)) && (x.Distribucion.HasValue) && (x.Distribucion.Value > 0)).ToList<ConfiguracionTramoOfertaBE>().ToCollection<ConfiguracionTramoOfertaBE>();
                            //Collection<ConfiguracionTramoOfertaBE> auxRellenos = listaConfiguracionTramo.Where(x => (x.CodTramo.Equals(item.CodTramo))).ToList<ConfiguracionTramoOfertaBE>().ToCollection<ConfiguracionTramoOfertaBE>();
                            if ((auxRellenos == null) || (auxRellenos.Count.Equals(0)))
                            {
                                listaTramosEliminar.Add(item);
                            }
                            else
                            {
                                listaTramosEliminar.Clear();
                            }
                        }
                    }
                }
                else
                {
                    limiteTramos = 30000;
                    foreach (TramoBE item in auxTramo.OrderBy(x => x.CodTramo))
                    {
                        if (item.CodTramoOrdenacion > limiteTramos)
                        {
                            //[FIX][MMUNOZ] Quieren que se muestren todos los valores de los tramos, tengan distribución o no
                            Collection<ConfiguracionTramoOfertaBE> auxRellenos = listaConfiguracionTramo.Where(x => (x.CodTramo.Equals(item.CodTramo)) && (x.Distribucion.HasValue) && (x.Distribucion.Value > 0)).ToList<ConfiguracionTramoOfertaBE>().ToCollection<ConfiguracionTramoOfertaBE>();
                            //Collection<ConfiguracionTramoOfertaBE> auxRellenos = listaConfiguracionTramo.Where(x => (x.CodTramo.Equals(item.CodTramo))).ToList<ConfiguracionTramoOfertaBE>().ToCollection<ConfiguracionTramoOfertaBE>();
                            if ((auxRellenos == null) || (auxRellenos.Count.Equals(0)))
                            {
                                listaTramosEliminar.Add(item);
                            }
                            else
                            {
                                listaTramosEliminar.Clear();
                            }
                        }
                    }
                }

                numTramos = auxTramo.Count;
                //Eliminamos los tramos que no deben insertarse
                foreach (TramoBE item in listaTramosEliminar)
                {
                    //auxTramo.Remove(item);
                    numTramos--;
                }

                matrizValores = new string[numTramos, numDestinos];
                int i = 0;
                int j = 0;

                //Para los productos S0134 y S0235 y los destinos Z7, Z8 y Z9 No se deben mostrar los tramos de expediciones
                //JCNS. MOSTRAR TRAMO. LO CAMBIO DE SITIO
                //bool DebeMostrarTramo = true;

                int maxDecimalesTarifa = 2;

                foreach (DestinoBE destino in listadoDestinos)
                {
                    j = 0;
                    foreach (TramoBE tramo in destino.Tramos)
                    {
                        //Para los productos S0134, S0132 y S0235 y los destinos Z7, Z8 y Z9 No se deben mostrar los tramos de expediciones
                        //JCNS. MOSTRAR TRAMO. LO CAMBIO DE SITIO
                        bool DebeMostrarTramo = true;

                        if ((objProductoBE.CodProducto.Equals("S0236") || objProductoBE.CodProducto.Equals("S0133") || objProductoBE.CodProducto.Equals("S0134") || objProductoBE.CodProducto.Equals("S0235") || objProductoBE.CodProducto.Equals("S0132")) && (destino.CodDestinoSAP.Equals("Z7") || destino.CodDestinoSAP.Equals("Z8") || destino.CodDestinoSAP.Equals("Z9")) && (tramo.CodTramo.StartsWith("E")))
                        {
                            DebeMostrarTramo = false;
                        }

                        TramoBE auxEliminar = listaTramosEliminar.FirstOrDefault(x => x.CodTramo.Equals(tramo.CodTramo));
                        if (auxEliminar == null)
                        {
                            TramoBE aux = auxTramo.FirstOrDefault(x => x.CodTramo.Equals(tramo.CodTramo));
                            if (aux != null)
                            {
                                string tarifaTramo = Math.Round(tramo.Tarifa, 5).ToString();

                                if (tarifasConDescuento)
                                {
                                    double auxTarifa = 0;

                                    //Según el tipo de modalidad se calcula la tarifa de una forma u otra.
                                    if (objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoDestino))
                                    {
                                        //Se obtiene la configuración del destino
                                        ConfiguracionDestinoOfertaBE configDestino = listaConfiguracionDestino.FirstOrDefault(x => x.idDestino.Equals(tramo.idDestino.Value));

                                        if (configDestino != null)
                                        {
                                            //Si se quiere mostrar la tarifa con los descuentos
                                            if (configDestino.DescuentoFinal.HasValue && double.TryParse(configDestino.DescuentoFinal.Value.ToString(), out auxTarifa))
                                            {
                                                auxTarifa = tramo.Tarifa - (tramo.Tarifa * auxTarifa / 100);
                                                tarifaTramo = Math.Round(auxTarifa, 5).ToString();
                                            }
                                        }
                                    }
                                    else
                                    {
                                        //Se obtiene la configuración del tramo
                                        ConfiguracionTramoOfertaBE configTramo = listaConfiguracionTramo.FirstOrDefault(x => x.idTramo.Equals(tramo.idTramo));

                                        if (configTramo != null)
                                        {
                                            //Si se quiere mostrar la tarifa con los descuentos
                                            if ((objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoTramo)) &&
                                                    configTramo.DescuentoFinal.HasValue && double.TryParse(configTramo.DescuentoFinal.Value.ToString(), out auxTarifa))
                                            {
                                                auxTarifa = tramo.Tarifa - (tramo.Tarifa * auxTarifa / 100);

                                                tarifaTramo = Math.Round(auxTarifa, 5).ToString();
                                            }
                                            else if ((objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoPrecioCierto)) &&
                                                    configTramo.PrecioCierto.HasValue && double.TryParse(configTramo.PrecioCierto.Value.ToString(), out auxTarifa))
                                            {
                                                if (!auxTarifa.Equals(0))
                                                {
                                                    tarifaTramo = Math.Round(auxTarifa, 5).ToString();
                                                    }
                                                    }
                                                }
                                            }
                                        }

                                if (DebeMostrarTramo)
                                {
                                    decimal argument = decimal.Parse(tarifaTramo);
                                    int count = BitConverter.GetBytes(decimal.GetBits(argument)[3])[2];

                                    if (count >= maxDecimalesTarifa)
                                        maxDecimalesTarifa = count;

                                    matrizValores[j, i] = tarifaTramo;
                                }
                                else
                                {
                                    matrizValores[j, i] = string.Empty;
                                }

                                j++;
                                DebeMostrarTramo = true;
                            }
                        }
                    }
                    i++;
                }

                //Asignamos el decimal más grande de los posibles
                for (i = 0; i < matrizValores.GetLength(0); i++)
                {
                    for (j = 0; j < matrizValores.GetLength(1); j++)
                    {
                        if (!String.IsNullOrEmpty(matrizValores[i, j]))
                        {
                            double tarifa = double.Parse(matrizValores[i, j]);
                            matrizValores[i, j] = tarifa.ToString("N" + maxDecimalesTarifa) + '€';
                        }
                    }
                }


                //Se buscan los datos de los VA            
                ListaTarifasVAReporte = this.CrearTablaVA(objProductoBE.idProducto, objProductoOfertaBE.idProductoOferta, tarifasConDescuento);

                //Se busca el valor máximo de peso volumétrico
                int maxPesoVolumetrico; //40  o 60

                //Vemos si se ha definido un peso volumétrico máximo especifico para el tipo de producto actual (CodAnexoSAP)
                maxPesoVolumetrico = new ProductoSAPBL().ObtenerPesoVolumetricoMaxProducto(objProductoBE.CodProducto);

                //Si no hay valor máximo registrado, se coge el valor por defecto (40 Kg).
                if (maxPesoVolumetrico == 0) maxPesoVolumetrico = 40;
                String textoPesoVolumetrico = string.Format(SimuladorResources.PesoVolumetricoMaximoInforme, maxPesoVolumetrico);

                #endregion

                #region Generacion Plantilla

                //Copiamos la plantilla en el fichero temporal.
                File.Copy(rutaPlantilla, ficheroTemporalWord, true);
                //Abrimos el fichero
                objWord.AbrirFichero();

                //INICIO CAMBIOS **********************************************

                int numTablasCrear = 1;
                int indiceTabla = 0;
                int numDestinosPorTabla = 15;

                //Paqueteria kg, si no en g. Miramos si su modelo de descuento es paquetería, tramos, o aparece como paquete en el doc. de tarifas   
                //Si es Publicorreo óptimo, no lo consideramos como paquetería     
                bool paqueteria = !ModeloDescuentoEnum.GetPaqueteriasQueSonPublicorreos().Contains(objProductoBE.CodProducto) && objProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Paqueteria)) ||
                                  objProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Tramos)) ||
                                  listaTramosInfTarifas.Count > 0;

                //Si es de paquetería internacional
                if (objProductoBE.Internacional)
                {
                    numTablasCrear = (int)Math.Ceiling(numDestinos / 15.0);

                    if (numTablasCrear != 1)
                    {
                        numDestinosPorTabla = 15;
                    }
                    else
                    {
                        numDestinosPorTabla = numDestinos;
                    }
                }
                //Si es normal
                else
                {
                    numTablasCrear = 1;
                    numDestinosPorTabla = numDestinos;
                }

                if (objWord.Abierto)
                {
                    objWord.ActivarModoWord();

                    #region Etiquetas word

                    if (numDestinosPorTabla <= 9)
                    {
                        objWord.CambiarOrientacionVertical();
                    }
                    else if (numDestinosPorTabla > 16)
                    {
                        objWord.AgrandarAnchuraWord(numDestinos - 16);
                    }

                    string textoParametro = string.Empty;
                    if (!tarifasConDescuento)
                    {
                        textoParametro = SimuladorResources.TituloReportTarifas;
                    }
                    else
                    {
                        textoParametro = SimuladorResources.TituloReportPrecios;
                    }

                    String saltoLinea;
                    
                    //Sustituimos las etiquetas,
                    objEtiquetas.Add("$$TITULOREPORTE", textoParametro);
                    objEtiquetas.Add("$$CODIGOSAP", string.Format(CultureInfo.InvariantCulture, "{0}", objProductoBE.Descripcion));
                    objEtiquetas.Add("$$VALIDEZDESDE", fechaInicial);
                    objEtiquetas.Add("$$VALIDEZHASTA", fechaFinal);
                    objEtiquetas.Add("$$NOMBRECLIENTECOMERCIAL", nombreCliente);
                    objEtiquetas.Add("$$TABLAGRTAMOS", string.Empty);
                    objEtiquetas.Add("$$MAXPESOVOLUMETRICO", textoPesoVolumetrico);                    
                    objWord.BuscaYSustituye(objEtiquetas, true, false);
                    #endregion

                    #region Rellenar las tablas


                    //List<DestinoBE> listadoDestinos = objProductoBE.Destinos.Where(x => x.Tramos.Count > 0).OrderBy(x => x.Orden).ToList();

                    //Duplicamos la primera tabla tantas veces como tablas a crear
                    for (int z = 0; z < numTablasCrear - 1; z++)
                    {

                        objWord.CopiarTabla(1);
                    }


                    int numDestinosPorTablaInicial = numDestinosPorTabla;

                    //Creamos T tablas, de acuerdo al tipo de producto
                    for (int iTabla = 0; iTabla < numTablasCrear; iTabla++)
                    {
                        indiceTabla++;
                        //Si es la última tabla, nos aseguramos del nº de columnas a mostrar    
                        if ((iTabla == numTablasCrear - 1) && (numDestinos % 15 != 0))
                        {
                            numDestinosPorTabla = numDestinos % 15;
                        }

                        objWord.SeleccionarTabla(iTabla + 1);

                        //Es Informe de Precios DD
                        var esInformePreciosDD = (objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoDestino) && tarifasConDescuento);

                    string tablaEnString = "\t"; // La primera celda de la cabecera es vacía
                        String descripciones = "\t";
                    
                        if (!esInformePreciosDD)
                    {
                            descripciones = "Peso\t";
                        }

                        for (int iDestino = 0; iDestino < numDestinosPorTabla; iDestino++)
                        {
                            DestinoBE objDestino = listadoDestinos[(iTabla * numDestinosPorTablaInicial) + iDestino];

                            //Obtenemos la descripcion
                            var objDescDestino = listaDestinos.FirstOrDefault(x => x.CodDestinoSAP.Equals(objDestino.CodDestinoSAPSinZona));
                            var descDestino = (objDescDestino != null && !String.IsNullOrEmpty(objDescDestino.DescripcionDestino)) ? objDescDestino.DescripcionDestino : "";
                            descDestino = StringUtil.ToTitleCaseIfAllUpper(descDestino);
                        // Añadimos cada destino
                            tablaEnString += objDestino.CodDestinoSAPSinZona + '\t'; // objDestino.CodDestinoSAP + descDestino + '\t';
                            descripciones += descDestino + '\t';
                    }

                    // Borramos el caracter extra insertado en el for
                    tablaEnString = tablaEnString.Substring(0, tablaEnString.Length - 1);
                    tablaEnString += '\n';
                        tablaEnString += descripciones.Substring(0, descripciones.Length - 1);
                        tablaEnString += '\n';

                        //Si se trata de descuento por destino, hay que añadir el descuento aplicado sobre tarifa
                        if (esInformePreciosDD)
                        {
                            tablaEnString += "Descuento aplicado sobre tarifa" + '\t';

                            for (int iDestino = 0; iDestino < numDestinosPorTabla; iDestino++)
                            {
                                DestinoBE objDestino = listadoDestinos[(iTabla * numDestinosPorTablaInicial) + iDestino];

                                //Obtenemos la configuración del destino con tramos
                                var configDestino = listaConfiguracionDestino.FirstOrDefault(t => t.idDestino.Equals(objDestino.idDestino));

                                //Si el destino tiene configuración y descuento final, mostramos el descuento en la tabla
                                if (configDestino != null && configDestino.DescuentoFinal.HasValue)
                                {
                                    tablaEnString += configDestino.DescuentoFinal.Value.ToString("N2") + "%\t";
                                }
                                else
                                {
                                    tablaEnString += "0.00%\t";
                                }
                            }

                            // Borramos el caracter extra insertado en el for
                            tablaEnString = tablaEnString.Substring(0, tablaEnString.Length - 1);
                            tablaEnString += '\n';
                            tablaEnString += "Peso\tPRECIO FINAL\n";
                        }

                    int fila = 0;
                        Decimal pesoTramo = 0;
                        String descripcionTramo = String.Empty;
                        String celda = "";

                        //Si es modalidad DD, o informe de tarifa de paquetería, mostramos los tramos del doc de tarifas
                        if (listaTramosInfTarifas.Count > 0 && (objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoDestino) || (paqueteria && !tarifasConDescuento)))
                        {
                            //Recorermos las listas de tarifas
                            foreach (TramoInformeBE tramoTarifas in listaTramosInfTarifas)
                            {
                                //Obtenemos los tramos inicio y fin del esquema de tramos de tarias
                                int tramoInicial = tramoTarifas.TramoIniGr.HasValue ? tramoTarifas.TramoIniGr.Value : int.MaxValue;
                                int tramoFinal = tramoTarifas.TramoFinGr.HasValue ? tramoTarifas.TramoFinGr.Value : int.MaxValue;

                                TramoBE objTramo = auxTramo.FirstOrDefault(t => t.RangoDeTramo.Item1 >= tramoInicial && t.RangoDeTramo.Item2 <= tramoFinal);
                                int indObjTramo = auxTramo.IndexOf(objTramo);

                                //Obtenemos la descripción del tramo                             
                                descripcionTramo = tramoTarifas.Descripcion;
                                tablaEnString += descripcionTramo + '\t';

                                //Si tiene * la descripción, mostraremos la aclaración * Para expediciones
                                if (objTramo != null && objTramo.CodTramo.Contains("E"))
                                    mostrarAclaracionExpediciones = true;

                                //Obtenemos el valor del precio/tarifa para el tramo
                                for (int columna = 0; columna < numDestinosPorTabla; columna++)
                                {
                                    celda = String.Empty;

                                    if (indObjTramo != -1)
                                    {
                                        if (tramoTarifas.TramoFinGr.HasValue)
                                        {
                                            //No es kg adicional
                                            celda = matrizValores[indObjTramo, (iTabla * numDestinosPorTablaInicial) + columna];
                                        }
                                        else
                                        {
                                            //Se trata de kg adicional: calculamos el valor por Kg. Adicional                                           
                                            celda = (double.Parse(matrizValores[indObjTramo + 1, (iTabla * numDestinosPorTablaInicial) + columna].Replace("€", "")) -
                                                    double.Parse(matrizValores[indObjTramo, (iTabla * numDestinosPorTablaInicial) + columna]
                                                    .Replace("€", ""))).ToString("N" + maxDecimalesTarifa) + "€";
                                        }
                                    }

                                    tablaEnString += celda + '\t';
                                }

                                // Borramos el caracter extra insertado en el for
                                tablaEnString = tablaEnString.Substring(0, tablaEnString.Length - 1);
                                tablaEnString += '\n';

                                fila++;

                                if (fila >= numTramos)
                                {
                                    break;
                                }
                            }
                        }
                        else
                        {
                            //foreach (TramoBE objTramo in auxTramo)
                            for (int t = 0; t < auxTramo.Count; t++)
                            {
                                TramoBE objTramo = auxTramo[t];
                                celda = "";

                        // Primera celda de cada tramo, contiene la descripción del tramo
                                descripcionTramo = String.Empty;

                                //Si es paquetería se muestra el texto Más De __ Kg hasta __ Kg
                                if (paqueteria && !objProductoOfertaBE.CodProductoSAP.Equals("S0360"))
                                {
                                    Func<decimal, string> formatearKg = (numDec => (numDec % 1) != 0 ? numDec.ToString("N3") : numDec.ToString("N3").TrimEnd('0').TrimEnd(','));

                                    if (t == 0)
                                    {
                                        pesoTramo = ((decimal)objTramo.RangoDeTramo.Item2) / 1000;
                                        descripcionTramo = "Hasta " + formatearKg(pesoTramo) + " Kg.";
                                    }
                                    else if (t < auxTramo.Count)// && objTramo.RangoDeTramo.Item2 != 30000)
                                    {
                                        descripcionTramo = "Más de " + formatearKg(pesoTramo) + " hasta ";
                                        pesoTramo = ((decimal)objTramo.RangoDeTramo.Item2) / 1000;
                                        descripcionTramo += formatearKg(pesoTramo) + " Kg.";
                                    }
                                }
                                //Si no es Paquetería, se muestra en Gramos.
                                else
                                {
                                    if (objTramo.RangoDeTramo.Item2.HasValue)
                                    {
                                        if (t == 0)
                                        {
                                            pesoTramo = ((decimal)objTramo.RangoDeTramo.Item2);
                                            descripcionTramo = "Hasta " + pesoTramo + " g.";
                                        }
                                        else if (t < auxTramo.Count)
                                        {
                                            descripcionTramo = "Más de " + pesoTramo + " hasta ";
                                            pesoTramo = ((decimal)objTramo.RangoDeTramo.Item2);
                                            descripcionTramo += pesoTramo + " g.";
                                        }
                                    }
                                    else
                                    {
                                        descripcionTramo += (decimal)objTramo.RangoDeTramo.Item1 + " g.";
                                    }

                                    //tablaEnString += (objTramo.Descripcion.Replace(" GRS", "")).Replace(".", "") + '\t';
                                }

                                if (objTramo.Descripcion.Contains("N"))
                                    descripcionTramo += " normalizadas";

                                //Si tiene * la descripción, mostraremos la aclaración * Para expediciones
                                if (objTramo.CodTramo.Contains("E"))
                                    mostrarAclaracionExpediciones = true;

                                tablaEnString += descripcionTramo + '\t';

                                for (int columna = 0; columna < numDestinosPorTabla; columna++)
                        {
                            // Valores en € del tramo
                                    celda = matrizValores[fila, (iTabla * numDestinosPorTablaInicial) + columna];
                            tablaEnString += celda + '\t';
                        }

                        // Borramos el caracter extra insertado en el for
                        tablaEnString = tablaEnString.Substring(0, tablaEnString.Length - 1);
                        tablaEnString += '\n';

                        fila++;

                        if (fila >= numTramos)
                        {
                            break;
                        }
                    }
                        }

                    // Borramos el caracter extra insertado en el for
                    tablaEnString = tablaEnString.Substring(0, tablaEnString.Length - 1);
                        objWord.EscribirTablaDeString(tablaEnString, true, maxDecimalesTarifa, esFormatoDesdeHasta: objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoDestino));
                    //[MMUNOZ] A partir de 60 columnas Word no puede controlar la tabla
                        if (numDestinos <= 60)
                        objWord.PonerTablaCabeceraRepetida();

                        //Si es DD o el informe de tarifas de un producto de paquetería, mostramos la tabla como DD
                        if (objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoDestino) || (paqueteria && !tarifasConDescuento))
                        {
                            objWord.DarFormatoTablaDescuentoDestino(esInformeTarifas: !tarifasConDescuento);
                        }
                        else if (objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoPrecioCierto) ||
                                objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoTramo))
                    {
                            objWord.DarFormatoTablaPrecioCierto();
                        }

                    }

                    //FIN CAMBIOS!!!!


                    if (mostrarAclaracionExpediciones)
                        objWord.EscribirTextoAlFinalDelDocumento(SimuladorResources.AclaracionParaExpediciones + "\n");

                    //if (objProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Tramos)) && mostrarAclaracionExpediciones)
                    if (mostrarAclaracionExpediciones)
                    {
                        if (objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoDestino))
                        {
                            objWord.EscribirTextoAlFinalDelDocumento(SimuladorResources.AclaracionKiloAdicional.Replace("60", "30") + "\n");
                        }
                        else
                    {
                            objWord.EscribirTextoAlFinalDelDocumento(SimuladorResources.AclaracionKiloAdicional + "\n");
                        }
                    }

                    //if (auxTramo.Count > 25)
                    //{
                    //    objWord.InsertarSaltoPaginaAlFinalDocumentoSiNecesario();
                    //}
                    objWord.EscribirTextoAlFinalDelDocumento("$$TABLAVA");

                    //Rellenamos la tabla de VA.
                    if (ListaTarifasVAReporte.Count > 0)
                    {
                        //insertamos la plantilla de la tabla de VA
                        objWord.InsertarDocumentoSustituyendoTexto(rutaPlantillaVA, "$$TABLAVA", (int)FormasSustitucionTexto.CopiarPegarConFormato);
                        int numPaginas6 = objWord.GetPagesNumber();
                        objWord.SeleccionarTabla(indiceTabla + 1);
                        string[] colsVA = new string[2];
                        //colsVA[0] = "LISTADO DE VALORES AÑADIDOS";
                        //colsVA[1] = string.Empty;
                        //objWord.AgregarTituloSinFormatoATabla(colsVA);
                        foreach (ReporteVABE objVA in ListaTarifasVAReporte)
                        {
                            colsVA[0] = objVA.Nombre;
                            colsVA[1] = objVA.Descripcion;
                            objWord.AgregarFilaSinFormatoATabla(colsVA);
                        }
                        //objWord.EliminarFilaDeTabla(2,2);
                        objWord.PonerTablaCabeceraRepetida();
                        objWord.DarFormatoTablaVA();

                    }
                    else
                    {
                        objWord.BuscaYSustituye("$$TABLAVA", "$$OBSERVACIONES", true, false);
                    }

                    #endregion

                    #region insertar pie de informe según linea de producto

                    int numPaginas = objWord.GetPagesNumber();
                    objWord.InsertarDocumentoSustituyendoTexto(rutaLineaProducto, "$$OBSERVACIONES", (int)FormasSustitucionTexto.CopiarPegarConFormato);
                    int numPaginasPost = objWord.GetPagesNumber();
                    saltoLinea = (numPaginasPost > numPaginas) ? "\f" : "";
                    objWord.BuscaYSustituye("$$SALTOLINEA", saltoLinea, true, false);                                      

                    //Insertamos dimensiones extra
                    //objWord.InsertarDocumentoSustituyendoTexto(rutaPlantillaDimExtra, "$$DIMENSIONESEXTRA", (int)FormasSustitucionTexto.CopiarPegarConFormato);

                    InformacionDestinosBL objInfoDestinosBL = new InformacionDestinosBL();
                    Collection<InformacionDestinosBE> objListaDestinos = objInfoDestinosBL.ObtenerListadoInformacionDestinos(objProductoBE.CodProducto);
                    StringBuilder sb = new StringBuilder();
                    if (objListaDestinos != null)
                    {
                        int insertados = 1;
                        foreach (DestinoBE objDestino in listadoDestinos)
                        {
                            InformacionDestinosBE objDescripcion = objListaDestinos.FirstOrDefault(x => x.CodDestinoSAP.Equals(objDestino.CodDestinoSAP));

                            if ((objDescripcion != null) && (!string.IsNullOrWhiteSpace(objDescripcion.DescripcionDestino)))
                            {
                                if (insertados < numDestinos)
                                {
                                    sb.Append(string.Format(CultureInfo.InvariantCulture, "{0}: {1}, ", objDescripcion.CodDestinoSAP, objDescripcion.DescripcionDestino));
                                }
                                else
                                {
                                    sb.Append(string.Format(CultureInfo.InvariantCulture, "{0}: {1}. ", objDescripcion.CodDestinoSAP, objDescripcion.DescripcionDestino));
                                }
                                insertados++;
                            }
                        }

                        //En caso de que la lista de destinos del producto y de la DB difieran.
                        if (sb.ToString() != string.Empty) sb.Replace(',', '.', sb.Length - 2, 1);
                    }
                    //objWord.BuscaYSustituye("$$OBSERVACIONESDESTINOS", sb.ToString(), true, false);
                    objWord.BuscaYSustituye(objEtiquetas, true, false);

                    if (informeMultiple && !esUltimoInforme)
                        //objWord.InsertarSaltoPaginaAlFinalDocumento();
                        objWord.InsertarCambioSeccionAlFinalDocumento();

                    #endregion

                    #region Generar WORD

                    //Ya tenemos generado el documento, ahora lo guardamos en docx.                    
                    ficheroDestino = string.Empty;

                    if (tarifasConDescuento)
                    {
                        ficheroDestino = Path.Combine(System.IO.Path.GetTempPath(), string.Format(CultureInfo.InvariantCulture, "{0}_{7}_{1}{2}{3}{4}{5}{6}.docx", "InformePrecios", DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(),
                            DateTime.Now.Day.ToString(), DateTime.Now.Hour.ToString(), DateTime.Now.Minute.ToString(), DateTime.Now.Second.ToString(), objProductoBE.CodProducto));
                    }
                    else
                    {
                        //ficheroDestino = Path.Combine(System.IO.Path.GetTempPath(), string.Format(CultureInfo.InvariantCulture, "{0}_{1}.docx", "InformeTarifas", objProductoBE.CodProducto));
                        ficheroDestino = Path.Combine(System.IO.Path.GetTempPath(), string.Format(CultureInfo.InvariantCulture, "{0}_{7}_{1}{2}{3}{4}{5}{6}.docx", "InformeTarifas", DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(),
                           DateTime.Now.Day.ToString(), DateTime.Now.Hour.ToString(), DateTime.Now.Minute.ToString(), DateTime.Now.Second.ToString(), objProductoBE.CodProducto));

                        //if (File.Exists(ficheroDestino))
                        //{
                            //si existe puede que esté abierto
                            //cerrar el proceso
                        //Process objOpenWord = Process.GetProcessesByName("winword.exe").FirstOrDefault(x => x.MainWindowTitle.Contains(string.Format(CultureInfo.InvariantCulture, "{0}_{1}.docx", "InformeTarifas", objProductoBE.CodProducto)));
                        //if (objOpenWord != null)
                        //{
                        //    objOpenWord.Kill();
                        //}
                        // }
                    }

                    //Guardamos el fichero
                    if (informeMultiple)
                    {
                        objWord.GuardarComo(ficheroTemporalPdf);
                    }
                    else
                    {
                        objWord.GuardarComo(ficheroDestino);
                    }

                    #endregion
                }

                #endregion

            }
            catch (Exception ex)
            {
                RegistrarAccionesSimulador.GuardarExcepcion(ex);
            }
            finally
            {
                if (objWord.Abierto)
                {
                    objWord.CerrarWord();
                }

                if (!informeMultiple)
                {
                    //Una vez creado el fichero en temporal se abre.
                    if (System.IO.File.Exists(ficheroDestino))
                    {
                        System.Diagnostics.Process proc = new System.Diagnostics.Process();
                        proc.EnableRaisingEvents = false;
                        proc.StartInfo.FileName = ficheroDestino;
                        proc.Start();
                    }
                }
            }
        }

        /// <summary>
        /// Genera el informe de tarifas/precios en formato Excel
        /// </summary>
        /// <param name="objProductoBE">productoBE del que se muestra el informe</param>
        /// <param name="objProductoOfertaBE">productoOfertaBE del que se muestra el informe</param>
        /// <param name="fechaInicial">fecha inicial de validez de la oferta</param>
        /// <param name="fechaFinal">fecha final de validez de la oferta</param>
        /// <param name="nombreCliente">Nombre del cliente al que pertenece la oferta</param>
        /// <param name="tarifasConDescuento">indica si queremos el informe de precios o el de tarifas</param>
        /// <param name="informeMultiple">Indica si se debe crear un informe orientado a la generación de uno múltiple (sin guardar PDF y guardando Words)</param>
        /// <returns>Devuelve la ruta del fichero donde se ha guardado el informe</returns>
        private void GenerarInformeEstandarExcel(ProductoBE objProductoBE, ProductoOfertaBE objProductoOfertaBE, string fechaInicial, string fechaFinal, string nombreCliente, bool tarifasConDescuento, bool informeMultiple, string ficheroTemporalExcel)
        {
            #region Variables

            //En caso de ser un producto con Destinos se rellena esta matriz de informacion
            object[,] matrizValores = null;
            bool esTipoPrecioCierto = false;

            //Listado de valores que se ingresan en la tabla de valores añadidos si corresponde
            Collection<ReporteVABE> ListaTarifasVAReporte = new Collection<ReporteVABE>();

            //Lista de etiquetas con su correspondiente valor que se sustituye en el documento word
            Dictionary<string, string> objEtiquetas = new Dictionary<string, string>();

            //Contiene la ruta de la plantilla que se usa para generar el reporte de tarifas
            string rutaPlantilla = string.Format(Utils.GetValorFromAppConfig(AppSettingsEnum.RutaPlantillaInformeTarifas), AppDomain.CurrentDomain.BaseDirectory);
            rutaPlantilla = rutaPlantilla.Substring(0, rutaPlantilla.Length - 4) + "xlsx";

            //Se activa para los casos que se quiere mostrar la aclaración *Para expediciones.
            Boolean mostrarAclaracionExpediciones = false;

            string rutaLineaProducto = string.Format(Utils.GetValorFromAppConfig(AppSettingsEnum.RutaPlantillaLineaProducto), AppDomain.CurrentDomain.BaseDirectory, objProductoBE.PlantillaInformeTarifasPrecios);


            rutaLineaProducto = rutaLineaProducto.Substring(0, rutaLineaProducto.Length - 4) + "xlsx";

            string rutaPlantillaVA = string.Format(Utils.GetValorFromAppConfig(AppSettingsEnum.RutaPlantillaInformeTarifasVA), AppDomain.CurrentDomain.BaseDirectory);

            // Si no es un informe múltiple o no tiene nombre, poner nombre por defecto
            if (ficheroTemporalExcel.Equals(string.Empty))// || !informeMultiple)
            {
                if (tarifasConDescuento)
                {
                    ficheroTemporalExcel = Path.Combine(System.IO.Path.GetTempPath(), string.Format(CultureInfo.InvariantCulture, "{0}_{7}_{1}{2}{3}{4}{5}{6}.xlsx", "InformePrecios", DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(),
                        DateTime.Now.Day.ToString(), DateTime.Now.Hour.ToString(), DateTime.Now.Minute.ToString(), DateTime.Now.Second.ToString(), objProductoBE.CodProducto));
                }
                else
                {
                    ficheroTemporalExcel = Path.Combine(System.IO.Path.GetTempPath(), string.Format(CultureInfo.InvariantCulture, "{0}_{7}_{1}{2}{3}{4}{5}{6}.xlsx", "InformeTarifas", DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(),
                        DateTime.Now.Day.ToString(), DateTime.Now.Hour.ToString(), DateTime.Now.Minute.ToString(), DateTime.Now.Second.ToString(), objProductoBE.CodProducto));
                    //ficheroTemporalExcel = Path.Combine(System.IO.Path.GetTempPath(), string.Format(CultureInfo.InvariantCulture, "{0}_{1}.xlsx", "InformeTarifas", objProductoBE.CodProducto));
                }
            }

            //instancia del objeto word con el que vamos  trabajar para generar el informe de tarifas
            ManagerExcel objExcel = new ManagerExcel(ficheroTemporalExcel, false);

            //Lista de tramos que usamos para, al generar el informe, mostrar la cabecera de los tramos.
            List<TramoBE> auxTramo = new List<TramoBE>();

            Collection<TramoBE> listaTramosEliminar = new Collection<TramoBE>();

            #endregion

            try
            {
                #region Obtencion Datos

                // Guardamos si es precio cierto o tipo descuento
                if (objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoPrecioCierto) || objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorPrecioCiertoGrupoTramo))
                {
                    esTipoPrecioCierto = true;
                }
                else
                {
                    esTipoPrecioCierto = false;
                }

                //Obtenemos los datos que ha ingresado el usuario.
                ConfiguracionProductosBL configProducto = new ConfiguracionProductosBL();
                Collection<ConfiguracionDestinoOfertaBE> listaConfiguracionDestino = configProducto.ObtenerConfiguracionDestinoOferta(objProductoOfertaBE.idProductoOferta);
                Collection<ConfiguracionTramoOfertaBE> listaConfiguracionTramo = configProducto.ObtenerConfiguracionTramoOferta(objProductoOfertaBE.idProductoOferta);
                Collection<InformacionDestinosBE> listaDestinos = new InformacionDestinosBL().ObtenerListadoInformacionDestinos(objProductoBE.CodProducto);
                List<TramoInformeBE> listaTramosInfTarifas = new TramoInformeBL().ObtenerTramosInformeProducto(objProductoBE.CodProducto);
                
                //Paqueteria kg, si no en g. Miramos si su modelo de descuento es paquetería, tramos, o aparece como paquete en el doc. de tarifas                        
                //Si es Publicorreo óptimo, no lo consideramos como paquetería     
                bool paqueteria = !ModeloDescuentoEnum.GetPaqueteriasQueSonPublicorreos().Contains(objProductoBE.CodProducto) && objProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Paqueteria)) ||
                                  objProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Tramos)) ||
                                  listaTramosInfTarifas.Count > 0;


                List<DestinoBE> listadoDestinos = objProductoBE.Destinos.Where(x => x.Tramos.Count > 0).OrderBy(t => t.Orden).ToList();
               
                if (objProductoBE.Internacional)
                {
                    listadoDestinos = (from d in listadoDestinos
                                       join c in listaConfiguracionDestino on d.idDestino equals c.idDestino
                                       where c.Distribucion.HasValue && c.Distribucion.Value > 0
                                       select d).ToList();
                }

                int numDestinos = listadoDestinos.Count();

                //Producto con destinos
                int numTramos = 0;
                //obtenemos el número máximo de tramos de entre los destinos. Ademas nos guardamos los tramos para la cabecera de fila de la tabla del informe. 
                foreach (DestinoBE destino in listadoDestinos)
                {
                    if (numTramos < destino.Tramos.Count)
                    {
                        numTramos = destino.Tramos.Count;
                        //auxTramo = destino.Tramos;
                        TramoBE[] copiaTramos = new TramoBE[numTramos];
                        destino.Tramos.CopyTo(copiaTramos, 0);
                        auxTramo = copiaTramos.ToList();
                    }
                }

                decimal limiteTramos;

                if (objProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Publicorreo)))
                {
                    limiteTramos = 250;
                    foreach (TramoBE item in auxTramo.OrderBy(x => x.CodTramo))
                    {
                        if (item.CodTramoDecimal > limiteTramos)
                        {
                            //[FIX][MMUNOZ] Quieren que se muestren todos los valores de los tramos, tengan distribución o no
                            Collection<ConfiguracionTramoOfertaBE> auxRellenos = listaConfiguracionTramo.Where(x => (x.CodTramo.Equals(item.CodTramo)) && (x.Distribucion.HasValue) && (x.Distribucion.Value > 0)).ToList<ConfiguracionTramoOfertaBE>().ToCollection<ConfiguracionTramoOfertaBE>();
                            //Collection<ConfiguracionTramoOfertaBE> auxRellenos = listaConfiguracionTramo.Where(x => (x.CodTramo.Equals(item.CodTramo))).ToList<ConfiguracionTramoOfertaBE>().ToCollection<ConfiguracionTramoOfertaBE>();
                            if ((auxRellenos == null) || (auxRellenos.Count.Equals(0)))
                            {
                                listaTramosEliminar.Add(item);
                            }
                            else
                            {
                                listaTramosEliminar.Clear();
                            }
                        }
                    }
                }
                else
                {
                    limiteTramos = 30000;
                    foreach (TramoBE item in auxTramo.OrderBy(x => x.CodTramo))
                    {
                        if (item.CodTramoOrdenacion > limiteTramos)
                        {
                            //[FIX][MMUNOZ] Quieren que se muestren todos los valores de los tramos, tengan distribución o no
                            Collection<ConfiguracionTramoOfertaBE> auxRellenos = listaConfiguracionTramo.Where(x => (x.CodTramo.Equals(item.CodTramo)) && (x.Distribucion.HasValue) && (x.Distribucion.Value > 0)).ToList<ConfiguracionTramoOfertaBE>().ToCollection<ConfiguracionTramoOfertaBE>();
                            //Collection<ConfiguracionTramoOfertaBE> auxRellenos = listaConfiguracionTramo.Where(x => (x.CodTramo.Equals(item.CodTramo))).ToList<ConfiguracionTramoOfertaBE>().ToCollection<ConfiguracionTramoOfertaBE>();
                            if ((auxRellenos == null) || (auxRellenos.Count.Equals(0)))
                            {
                                listaTramosEliminar.Add(item);
                            }
                            else
                            {
                                listaTramosEliminar.Clear();
                            }
                        }
                    }
                }

                numTramos = auxTramo.Count;
                //Eliminamos los tramos que no deben insertarse
                foreach (TramoBE item in listaTramosEliminar)
                {
                    //auxTramo.Remove(item);
                    numTramos--;
                }

                matrizValores = new object[numTramos, numDestinos];
                int i = 0;
                int j = 0;
                decimal maxDecimalesTarifa = 2;

                //Para los productos S0134 y S0235 y los destinos Z7, Z8 y Z9 No se deben mostrar los tramos de expediciones
                //JCNS. MOSTRAR TRAMO. LO CAMBIO DE SITIO
                //bool DebeMostrarTramo = true;

                foreach (DestinoBE destino in listadoDestinos)
                {
                    j = 0;
                    foreach (TramoBE tramo in destino.Tramos)
                    {
                        //Para los productos S0134, S0132, S0235 S0133, o S0236 y los destinos Z7, Z8 y Z9 No se deben mostrar los tramos de expediciones
                        //JCNS. MOSTRAR TRAMO. LO CAMBIO DE SITIO
                        bool DebeMostrarTramo = true;

                        if ((objProductoBE.CodProducto.Equals("S0236") || objProductoBE.CodProducto.Equals("S0133") || objProductoBE.CodProducto.Equals("S0134") || objProductoBE.CodProducto.Equals("S0235") || objProductoBE.CodProducto.Equals("S0132")) && (destino.CodDestinoSAP.Equals("Z7") || destino.CodDestinoSAP.Equals("Z8") || destino.CodDestinoSAP.Equals("Z9")) && (tramo.CodTramo.StartsWith("E")))
                        {
                            DebeMostrarTramo = false;
                        }

                        TramoBE auxEliminar = listaTramosEliminar.FirstOrDefault(x => x.CodTramo.Equals(tramo.CodTramo));
                        if (auxEliminar == null)
                        {
                            TramoBE aux = auxTramo.FirstOrDefault(x => x.CodTramo.Equals(tramo.CodTramo));
                            if (aux != null)
                            {
                                object tarifaTramo = Math.Round(tramo.Tarifa, 5);

                                if (esTipoPrecioCierto)
                                {
                                    tarifaTramo = Math.Round(tramo.Tarifa, 5);
                                }

                                if (tarifasConDescuento)
                                {
                                    double auxTarifa = 0;

                                    //Según el tipo de modalidad se calcula la tarifa de una forma u otra.
                                    if (objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoDestino))
                                    {
                                        //Se obtiene la configuración del destino
                                        ConfiguracionDestinoOfertaBE configDestino = listaConfiguracionDestino.FirstOrDefault(x => x.idDestino.Equals(tramo.idDestino.Value));

                                        if (configDestino != null)
                                        {
                                            //Si se quiere mostrar la tarifa con los descuentos
                                            if (configDestino.DescuentoFinal.HasValue && double.TryParse(configDestino.DescuentoFinal.Value.ToString(), out auxTarifa))
                                            {
                                                auxTarifa = tramo.Tarifa - (tramo.Tarifa * auxTarifa / 100);
                                                tarifaTramo = Math.Round(auxTarifa, 5);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        //Se obtiene la configuración del tramo
                                        ConfiguracionTramoOfertaBE configTramo = listaConfiguracionTramo.FirstOrDefault(x => x.idTramo.Equals(tramo.idTramo));

                                        if (configTramo != null)
                                        {
                                            //Si se quiere mostrar la tarifa con los descuentos
                                            if ((objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoTramo)) &&
                                                    configTramo.DescuentoFinal.HasValue && double.TryParse(configTramo.DescuentoFinal.Value.ToString(), out auxTarifa))
                                            {
                                                auxTarifa = tramo.Tarifa - (tramo.Tarifa * auxTarifa / 100);

                                                tarifaTramo = Math.Round(auxTarifa, 5);
                                            }
                                            else if ((objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoPrecioCierto)) &&
                                                    configTramo.PrecioCierto.HasValue && double.TryParse(configTramo.PrecioCierto.Value.ToString(), out auxTarifa))
                                            {
                                                if (!auxTarifa.Equals(0))
                                                {
                                                    tarifaTramo = Math.Round(auxTarifa, 5);
                                                }
                                            }
                                        }
                                    }
                                }

                                if (DebeMostrarTramo)
                                {
                                    decimal argument = (decimal)(double)tarifaTramo;
                                    int count = BitConverter.GetBytes(decimal.GetBits(argument)[3])[2];

                                    if (count >= maxDecimalesTarifa)
                                        maxDecimalesTarifa = count;

                                    matrizValores[j, i] = tarifaTramo;
                                }
                                else
                                {
                                    matrizValores[j, i] = null;
                                }

                                j++;
                                DebeMostrarTramo = true;
                            }
                        }
                    }
                    i++;
                }

                //Se buscan los datos de los VA            
                ListaTarifasVAReporte = this.CrearTablaVA(objProductoBE.idProducto, objProductoOfertaBE.idProductoOferta, tarifasConDescuento);

                //Se busca el valor máximo de peso volumétrico
                int maxPesoVolumetrico; //40  o 60

                //Vemos si se ha definido un peso volumétrico máximo especifico para el tipo de producto actual (CodAnexoSAP)
                maxPesoVolumetrico = new ProductoSAPBL().ObtenerPesoVolumetricoMaxProducto(objProductoBE.CodProducto);

                //Si no hay valor máximo registrado, se coge el valor por defecto (40 Kg).
                if (maxPesoVolumetrico == 0) maxPesoVolumetrico = 40;
                String textoPesoVolumetrico = string.Format(SimuladorResources.PesoVolumetricoMaximoInforme, maxPesoVolumetrico);

                #endregion

                #region Generacion Plantilla

                bool esNecesarioInsertarPlantilla = true;
                //if (!informeMultiple || !File.Exists(ficheroTemporalExcel))
                if (!File.Exists(ficheroTemporalExcel))
                {
                    esNecesarioInsertarPlantilla = false;
                    //Copiamos la plantilla en el fichero temporal.
                    File.Copy(rutaPlantilla, ficheroTemporalExcel, true);
                }

                esNecesarioInsertarPlantilla = true;

                //Abrimos el fichero
                try
                {
                    objExcel.AbrirFichero();
                }
                catch { }
                if (objExcel.Abierto)
                {
                    objExcel.SeleccionarPrimeraHojaLibreOCrearNuevaYRenombrar(objProductoBE.CodAnexoSAP + " - " + objProductoBE.CodProducto + " - " + GenerarAbreviaturaModeloNegociacion(objProductoOfertaBE.CodModalidadNegociacion.Trim()));

                    #region Etiquetas excel

                    string textoParametro = string.Empty;
                    if (!tarifasConDescuento)
                    {
                        textoParametro = SimuladorResources.TituloReportTarifas;
                    }
                    else
                    {
                        textoParametro = SimuladorResources.TituloReportPrecios;
                    }

                    #endregion

                    #region Rellenar las tablas
                    int numTablasCrear = 1;
                    int numDestinosPorTabla = 15;

                    //Paqueteria kg, si no en g. Miramos si su modelo de descuento es paquetería, tramos, o aparece como paquete en el doc. de tarifas   
                    //Si es Publicorreo óptimo, no lo consideramos como paquetería     
                    //Si es de paquetería internacional
                    if (objProductoBE.Internacional)
                    {
                        numTablasCrear = (int)Math.Ceiling(numDestinos / 15.0);

                        if (numTablasCrear != 1)
                        {
                            numDestinosPorTabla = 15;
                        }
                        else
                        {
                            numDestinosPorTabla = numDestinos;
                        }
                    }
                    //Si es normal
                    else
                    {
                        numTablasCrear = 1;
                        numDestinosPorTabla = numDestinos;
                    }
                    

                    //Relleno la tabla como si hubiera destinos
                    object[] colsDestino, colsDescDestino;


                    int numDestinosPorTablaInicial = numDestinosPorTabla;

                    //Creamos T tablas, de acuerdo al tipo de producto
                    for (int iTabla = 0; iTabla < numTablasCrear; iTabla++)
                    {
                        //Si es la última tabla, nos aseguramos del nº de columnas a mostrar    
                        if ((iTabla == numTablasCrear - 1) && (numDestinos % 15 != 0))
                        {
                            numDestinosPorTabla = numDestinos % 15;
                        }

                    //añadir primera fila con las cabeceras de los destinos                             
                        colsDestino = new object[numDestinosPorTabla + 1];
                        colsDescDestino = new object[numDestinosPorTabla + 1];
                    colsDestino[0] = string.Empty;
                        colsDescDestino[0] = string.Empty;
                        
                        //Es Informe de Precios DD
                        var esInformePreciosDD = (objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoDestino) && tarifasConDescuento);

                        if (!esInformePreciosDD)
                        {
                            colsDescDestino[0] = "Peso";
                        }

                    int destino = 1;
                        for (int iDestino = 0; iDestino < numDestinosPorTabla; iDestino++)
                    {
                            DestinoBE objDestino = listadoDestinos[(iTabla * numDestinosPorTablaInicial) + iDestino];

                            //Obtenemos la descripcion
                            var objDescDestino = listaDestinos.FirstOrDefault(d => d.CodDestinoSAP.Equals(objDestino.CodDestinoSAPSinZona));
                            var descDestino = (objDescDestino != null && !String.IsNullOrEmpty(objDescDestino.DescripcionDestino)) ? objDescDestino.DescripcionDestino : String.Empty;
                            descDestino = StringUtil.ToTitleCaseIfAllUpper(descDestino);

                            colsDestino[destino] = objDestino.CodDestinoSAPSinZona;
                            colsDescDestino[destino] = descDestino;
                        destino++;
                    }

                    objExcel.IniciarDibujarTabla(true, true);

                        objExcel.AgregarFila(colsDestino, esAnchoFijo: true);
                        objExcel.AgregarFila(colsDescDestino, esAnchoFijo: true);
                        colsDestino = new object[numDestinosPorTabla + 1];


                        //Si se trata de descuento por destino, hay que añadir el descuento aplicado sobre tarifa
                        if (objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoDestino) && tarifasConDescuento)
                        {
                            int contadorCols = 1;
                            colsDestino[0] += "Descuento aplicado sobre tarifa" + '\t';
                            
                            for (int iDestino = 0; iDestino < numDestinosPorTabla; iDestino++)
                            {
                                DestinoBE objDestino = listadoDestinos[(iTabla * numDestinosPorTablaInicial) + iDestino];

                                //Obtenemos la configuración del destino con tramos
                                var configDestino = listaConfiguracionDestino.FirstOrDefault(t => t.idDestino.Equals(objDestino.idDestino));

                                //Si el destino tiene configuración y descuento final, mostramos el descuento en la tabla
                                if (configDestino != null && configDestino.DescuentoFinal.HasValue)
                                {
                                    colsDestino[contadorCols] = configDestino.DescuentoFinal.Value.ToString("N2") + "%";
                                }
                                else
                                {
                                    colsDestino[contadorCols] = "0,00%";
                                }

                                contadorCols++;
                            }
                            
                    objExcel.AgregarFila(colsDestino);
                            objExcel.AgregarFila(new String[] { "Peso", "PRECIO FINAL" });
                        }

                    //variable que uso para meter los valores en la plantilla.tiene numDestinos+1 para meter en la primera casilla                            
                    int fila = 0;
                        decimal pesoTramo = 0;
                        
                        //Ponemos los tramos del doc de tarifas si es DD, o si es paqueteria y doc de tarifas
                        if (listaTramosInfTarifas.Count > 0 && ((objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoDestino)) || (paqueteria && !tarifasConDescuento)))
                        {
                            //Recorermos las listas de tarifas
                            foreach (TramoInformeBE tramoTarifas in listaTramosInfTarifas)
                            {
                                //Obtenemos los tramos inicio y fin del esquema de tramos de tarias
                                int tramoInicial = tramoTarifas.TramoIniGr.HasValue ? tramoTarifas.TramoIniGr.Value : int.MaxValue;
                                int tramoFinal = tramoTarifas.TramoFinGr.HasValue ? tramoTarifas.TramoFinGr.Value : int.MaxValue;

                                TramoBE objTramo = auxTramo.FirstOrDefault(t => t.RangoDeTramo.Item1 >= tramoInicial && t.RangoDeTramo.Item2 <= tramoFinal);
                                int indObjTramo = auxTramo.IndexOf(objTramo);

                                //Obtenemos la descripción del tramo
                                colsDestino = new object[numDestinosPorTabla + 1];
                                colsDestino[0] = tramoTarifas.Descripcion;

                                if (objTramo != null && objTramo.CodTramo.Contains('E'))
                                    mostrarAclaracionExpediciones = true;

                                //Obtenemos el valor del precio/tarifa para el tramo
                                if (indObjTramo != -1)
                                {
                                    for (int columna = 0; columna < numDestinosPorTabla; columna++)
                                    {
                                        if (tramoTarifas.TramoFinGr.HasValue)
                                        {
                                            //No es kg adicional
                                            colsDestino[columna + 1] = matrizValores[indObjTramo, (iTabla * numDestinosPorTablaInicial) + columna] ?? 0.0;
                                        }
                                        else
                                        {
                                            //Se trata de kg adicional: calculamos el valor por Kg. Adicional
                                            colsDestino[columna + 1] = ((double)(matrizValores[indObjTramo + 1, (iTabla * numDestinosPorTablaInicial) + columna] ?? 0.0)) - ((double)(matrizValores[indObjTramo, (iTabla * numDestinosPorTablaInicial) + columna] ?? 0.0));
                                        }
                                    }
                                }

                                objExcel.AgregarFila(colsDestino, esAnchoFijo: true);
                                fila++;

                                if (fila >= numTramos)
                                {
                                    break;
                                }
                            }
                        }
                        else
                        {
                            for (int t = 0; t < auxTramo.Count; t++)
                            {
                                TramoBE objTramo = auxTramo[t];
                                colsDestino = new object[numDestinosPorTabla + 1];
                                
                                if (objTramo != null && objTramo.CodTramo.Contains('E'))
                                    mostrarAclaracionExpediciones = true;

                                if (paqueteria && !objProductoOfertaBE.CodProductoSAP.Equals("S0360"))
                                {

                                    Func<decimal, string> formatearKg = (numDec => (numDec % 1) != 0 ? numDec.ToString("N3") : numDec.ToString("N3").TrimEnd('0').TrimEnd(','));

                                    if (t == 0)
                                    {
                                        pesoTramo = ((decimal)objTramo.RangoDeTramo.Item2) / 1000;
                                        colsDestino[0] = "Hasta " + formatearKg(pesoTramo) + " Kg.";
                                    }
                                    else if (t < auxTramo.Count)// && objTramo.RangoDeTramo.Item2 != 30000)
                                    {
                                        colsDestino[0] = "Más de " + formatearKg(pesoTramo) + " hasta ";
                                        pesoTramo = ((decimal)objTramo.RangoDeTramo.Item2) / 1000;
                                        colsDestino[0] += formatearKg(pesoTramo) + " Kg.";
                                    }                                  
                                }
                                else
                                {

                                    if (objTramo.RangoDeTramo.Item2.HasValue)
                                    {

                                        if (t == 0)
                                        {
                                            pesoTramo = ((decimal)objTramo.RangoDeTramo.Item2);
                                            colsDestino[0] = "Hasta " + pesoTramo + " g.";
                                        }
                                        else if (t < auxTramo.Count)
                                        {
                                            colsDestino[0] = "Más de " + pesoTramo + " hasta ";
                                            pesoTramo = ((decimal)objTramo.RangoDeTramo.Item2);
                                            colsDestino[0] += pesoTramo + " g.";
                                        }
                                    }
                                    else
                    {
                                        colsDestino[0] = ((decimal)objTramo.RangoDeTramo.Item1) + " g";
                                    }
                                }

                                if (objTramo.Descripcion.Contains("N"))
                                    colsDestino[0] += " normalizadas";

                                for (int columna = 0; columna < numDestinosPorTabla; columna++)
                        {
                                    colsDestino[columna + 1] = matrizValores[fila, (iTabla * numDestinosPorTablaInicial) + columna];
                        }
                        objExcel.AgregarFila(colsDestino);
                        fila++;

                        if (fila >= numTramos)
                        {
                            break;
                        }

                            }
                        }

                        objExcel.TerminarDibujarTabla(false, false, (int?)maxDecimalesTarifa);

                        //Modificamos el formato de la tabla para la modalidad descuento por destino (o si es un producto de paquetería valo modalidad tarifa)
                        if (objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoDestino) || (paqueteria && !tarifasConDescuento))
                        {
                            objExcel.DarFormatoTablaDescuentoDestino(numDestinosPorTabla, esInformeTarifas: !tarifasConDescuento);
                        }
                        else if (objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoPrecioCierto) ||
                                 objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoTramo))
                        {
                            objExcel.DarFormatoTablaPrecioCierto(numDestinosPorTabla, numTramos);
                        }                       
                     
                    }

                    if (mostrarAclaracionExpediciones)
                    {
                        objExcel.EscribirTextoAlFinalDelDocumento(SimuladorResources.AclaracionParaExpediciones);
                    }

                    if (mostrarAclaracionExpediciones)
                    {
                        if (objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoDestino))
                        {
                            objExcel.EscribirTextoAlFinalDelDocumento(SimuladorResources.AclaracionKiloAdicional.Replace("60", "30"));
                        }
                        else
                    {
                        objExcel.EscribirTextoAlFinalDelDocumento(SimuladorResources.AclaracionKiloAdicional);
                    }
                    }

                    //Rellenamos la tabla de VA.
                    if (ListaTarifasVAReporte.Count > 0)
                    {
                        objExcel.IniciarDibujarTabla(true);
                        string[] colsVATitulo = new string[2];
                        colsVATitulo[0] = "LISTADO DE VALORES AÑADIDOS";
                        colsVATitulo[1] = string.Empty;
                        objExcel.AgregarFila(colsVATitulo, ajustarTextoColumna: false);
                        foreach (ReporteVABE objVA in ListaTarifasVAReporte)
                        {
                            string[] colsVA = new string[2];
                            colsVA[0] = objVA.Nombre;
                            colsVA[1] = objVA.Descripcion;
                            objExcel.AgregarFila(colsVA);
                        }
                        objExcel.TerminarDibujarTabla(false, false, null, true);
                    }

                    #endregion

                    #region Sustitución de tags

                    // Al contrario que en word, hacemos esto lo último para no tener que insertar la tabla entre medias
                    objExcel.InsertarDocumentoAlFinal(rutaLineaProducto);

                    //Sustituimos las etiquetas,
                    objEtiquetas.Add("$$TITULOREPORTE", textoParametro);
                    objEtiquetas.Add("$$CODIGOSAP", string.Format(CultureInfo.InvariantCulture, "{0}", objProductoBE.Descripcion));
                    objEtiquetas.Add("$$VALIDEZDESDE", fechaInicial);
                    objEtiquetas.Add("$$VALIDEZHASTA", fechaFinal);
                    objEtiquetas.Add("$$NOMBRECLIENTECOMERCIAL", nombreCliente);
                    objEtiquetas.Add("$$TABLAGRTAMOS", string.Empty);
                    objEtiquetas.Add("$$MAXPESOVOLUMETRICO", textoPesoVolumetrico);
                    objExcel.BuscaYSustituye(objEtiquetas, true, false);

                    #endregion

                    #region insertar pie de informe según linea de producto

                    InformacionDestinosBL objInfoDestinosBL = new InformacionDestinosBL();
                    Collection<InformacionDestinosBE> objListaDestinos = objInfoDestinosBL.ObtenerListadoInformacionDestinos(objProductoBE.CodProducto);
                    StringBuilder sb = new StringBuilder();
                    if (objListaDestinos != null)
                    {
                        int insertados = 1;
                        foreach (DestinoBE objDestino in listadoDestinos)
                        {
                            InformacionDestinosBE objDescripcion = objListaDestinos.FirstOrDefault(x => x.CodDestinoSAP.Equals(objDestino.CodDestinoSAP));

                            if ((objDescripcion != null) && (!string.IsNullOrWhiteSpace(objDescripcion.DescripcionDestino)))
                            {
                                if (insertados < numDestinos)
                                {
                                    sb.Append(string.Format(CultureInfo.InvariantCulture, "{0}: {1}, ", objDescripcion.CodDestinoSAP, objDescripcion.DescripcionDestino));
                                }
                                else
                                {
                                    sb.Append(string.Format(CultureInfo.InvariantCulture, "{0}: {1}. ", objDescripcion.CodDestinoSAP, objDescripcion.DescripcionDestino));
                                }
                                insertados++;
                            }
                        }

                        //En caso de que la lista de destinos del producto y de la DB difieran.
                        if (sb.ToString() != string.Empty) sb.Replace(',', '.', sb.Length - 2, 1);
                    }

                    //objExcel.BuscaYSustituye("$$OBSERVACIONESDESTINOS", sb.ToString(), true, false);
                    objExcel.BuscaYSustituye(objEtiquetas, true, false);

                    #endregion
                }

                #endregion

                #region Abrir el fichero

                if (objExcel.Abierto)
                {
                    objExcel.GuardarLibro();
                    objExcel.CerrarExcel();
                }

                if (!informeMultiple)
                {
                    ManagerExcel.AbrirExcelStandalone(ficheroTemporalExcel);
                }

                #endregion
            }
            catch (Exception ex)
            {
                RegistrarAccionesSimulador.GuardarExcepcion(ex);
            }
            finally
            {
                if (objExcel.Abierto)
                {
                    objExcel.CerrarExcel();
                }
            }
        }

        /// <summary>
        /// genera un informe de tarifas/precios en formato pdf con modo de publicorreo
        /// </summary>
        /// <param name="objProductoBE"></param>
        /// <param name="objProductoOfertaBE"></param>
        /// <param name="fechaInicial"></param>
        /// <param name="fechaFinal"></param>
        /// <param name="nombreCliente"></param>
        /// <param name="tarifasConDescuento"></param>
        private void GenerarInformeEstandarPubliCorreo(ProductoBE objProductoBE, ProductoOfertaBE objProductoOfertaBE, string fechaInicial, string fechaFinal, string nombreCliente, bool tarifasConDescuento, bool informeMultiple, string ficheroTemporalWord, string ficheroTemporalPdf, bool esUltimoInforme)
        {
            #region Variables

            //En caso de ser un producto con Destinos se rellena esta matriz de informacion
            string[,] matrizValores = null;

            //Listado de valores que se ingresan en la tabla de valores añadidos si corresponde
            Collection<ReporteVABE> ListaTarifasVAReporte = new Collection<ReporteVABE>();

            //Lista de etiquetas con su correspondiente valor que se sustituye en el documento word
            Dictionary<string, object> objEtiquetas = new Dictionary<string, object>();

            //Contiene la ruta de la plantilla que se usa para generar el reporte de tarifas
            string rutaPlantilla = string.Format(Utils.GetValorFromAppConfig(AppSettingsEnum.RutaPlantillaInformeTarifasPubliCorreo), AppDomain.CurrentDomain.BaseDirectory);

            string rutaLineaProducto = string.Format(Utils.GetValorFromAppConfig(AppSettingsEnum.RutaPlantillaLineaProducto), AppDomain.CurrentDomain.BaseDirectory, objProductoBE.PlantillaInformeTarifasPrecios);
            string rutaPlantillaVA = string.Format(Utils.GetValorFromAppConfig(AppSettingsEnum.RutaPlantillaInformeTarifasVA), AppDomain.CurrentDomain.BaseDirectory);

            // Si no es un informe múltiple o no tiene nombre, poner nombre por defecto
            if (ficheroTemporalWord.Equals(string.Empty) || !informeMultiple)
            {
                ficheroTemporalWord = Path.Combine(System.IO.Path.GetTempPath(), "tarifas.docx");
            }

            string ficheroDestino = string.Empty;

            //instancia del objeto word con el que vamos  trabajar para generar el informe de tarifas
            ManagerWord objWord = new ManagerWord(ficheroTemporalWord, false);

            //Lista de tramos que usamos para, al generar el informe, mostrar la cabecera de los tramos.
            List<TramoBE> auxTramo = new List<TramoBE>();

            Collection<TramoBE> listaTramosEliminar = new Collection<TramoBE>();

            #endregion

            try
            {
                #region Obtencion Datos

                //Obtenemos los datos que ha ingresado el usuario.
                ConfiguracionProductosBL configProducto = new ConfiguracionProductosBL();
                Collection<ConfiguracionDestinoOfertaBE> listaConfiguracionDestino = configProducto.ObtenerConfiguracionDestinoOferta(objProductoOfertaBE.idProductoOferta);
                Collection<ConfiguracionTramoOfertaBE> listaConfiguracionTramo = configProducto.ObtenerConfiguracionTramoOferta(objProductoOfertaBE.idProductoOferta);
                Collection<InformacionDestinosBE> listaDestinos = new InformacionDestinosBL().ObtenerListadoInformacionDestinos(objProductoBE.CodProducto);

                var listadoDestinos = objProductoBE.Destinos.Where(x => x.Tramos.Count > 0).OrderBy(t => t.Orden).ToList();

                if (objProductoBE.Internacional)
                {
                    listadoDestinos = (from d in listadoDestinos
                                       join c in listaConfiguracionDestino on d.idDestino equals c.idDestino
                                       where c.Distribucion.HasValue && c.Distribucion.Value > 0
                                       select d).ToList();

                }

                int numDestinos = listadoDestinos.Count();


                //Producto con destinos
                int numTramos = 0;
                //obtenemos el número máximo de tramos de entre los destinos. Ademas nos guardamos los tramos para la cabecera de fila de la tabla del informe. 
                foreach (DestinoBE destino in listadoDestinos)
                {
                    if (numTramos < destino.Tramos.Count)
                    {
                        numTramos = destino.Tramos.Count;
                        //auxTramo = destino.Tramos;
                        TramoBE[] copiaTramos = new TramoBE[numTramos];
                        destino.Tramos.CopyTo(copiaTramos, 0);
                        auxTramo = copiaTramos.ToList();
                    }
                }

                decimal limiteTramos = 250;

                foreach (TramoBE item in auxTramo.OrderBy(x => x.CodTramo))
                {
                    if (item.CodTramoDecimal > limiteTramos)
                    {
                        //[FIX][MMUNOZ] Quieren que se muestren todos los valores de los tramos, tengan distribución o no
                        Collection<ConfiguracionTramoOfertaBE> auxRellenos = listaConfiguracionTramo.Where(x => (x.CodTramo.Equals(item.CodTramo)) && (x.Distribucion.HasValue) && (x.Distribucion.Value > 0)).ToList<ConfiguracionTramoOfertaBE>().ToCollection<ConfiguracionTramoOfertaBE>();
                        
                        if ((auxRellenos == null) || (auxRellenos.Count.Equals(0)))
                        {
                            listaTramosEliminar.Add(item);
                        }
                        else
                        {
                            listaTramosEliminar.Clear();
                        }
                    }
                }

                numTramos = auxTramo.Count;
                //Eliminamos los tramos que no deben insertarse
                foreach (TramoBE item in listaTramosEliminar)
                {
                    //auxTramo.Remove(item);
                    numTramos--;
                }

                matrizValores = new string[numTramos, numDestinos];
                int i = 0;
                int j = 0;
                int maxDecimalesTarifa = 2;

                foreach (DestinoBE destino in listadoDestinos)
                {
                    j = 0;
                    foreach (TramoBE tramo in destino.Tramos)
                    {
                        TramoBE auxEliminar = listaTramosEliminar.FirstOrDefault(x => x.CodTramo.Equals(tramo.CodTramo));
                        if (auxEliminar == null)
                        {
                            TramoBE aux = auxTramo.FirstOrDefault(x => x.CodTramo.Equals(tramo.CodTramo));
                            if (aux != null)
                            {
                                string tarifaTramo = Math.Round(tramo.Tarifa, 5).ToString();

                                if (tarifasConDescuento)
                                {
                                    double auxTarifa = 0;

                                    //Según el tipo de modalidad se calcula la tarifa de una forma u otra.
                                    if (objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoDestino))
                                    {
                                        //Se obtiene la configuración del destino
                                        ConfiguracionDestinoOfertaBE configDestino = listaConfiguracionDestino.FirstOrDefault(x => x.idDestino.Equals(tramo.idDestino.Value));

                                        if (configDestino != null)
                                        {
                                            //Si se quiere mostrar la tarifa con los descuentos
                                            if (configDestino.DescuentoFinal.HasValue && double.TryParse(configDestino.DescuentoFinal.Value.ToString(), out auxTarifa))
                                            {
                                                auxTarifa = tramo.Tarifa - (tramo.Tarifa * auxTarifa / 100);
                                                tarifaTramo = Math.Round(auxTarifa, 5).ToString();
                                            }
                                        }
                                    }
                                    else
                                    {
                                        //Se obtiene la configuración del tramo
                                        ConfiguracionTramoOfertaBE configTramo = listaConfiguracionTramo.FirstOrDefault(x => x.idTramo.Equals(tramo.idTramo));

                                        if (configTramo != null)
                                        {
                                            //Si se quiere mostrar la tarifa con los descuentos
                                            if ((objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoTramo)) &&
                                                    configTramo.DescuentoFinal.HasValue && double.TryParse(configTramo.DescuentoFinal.Value.ToString(), out auxTarifa))
                                            {
                                                auxTarifa = tramo.Tarifa - (tramo.Tarifa * auxTarifa / 100);

                                                tarifaTramo = Math.Round(auxTarifa, 5).ToString();
                                            }
                                            else if ((objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoPrecioCierto)) &&
                                                    configTramo.PrecioCierto.HasValue && double.TryParse(configTramo.PrecioCierto.Value.ToString(), out auxTarifa))
                                            {
                                                if (!auxTarifa.Equals(0))
                                                {
                                                        tarifaTramo = Math.Round(auxTarifa, 5).ToString();
                                                    }
                                            }
                                        }
                                    }
                                }

                                decimal argument = decimal.Parse(tarifaTramo);
                                int count = BitConverter.GetBytes(decimal.GetBits(argument)[3])[2];

                                if (count >= maxDecimalesTarifa)
                                    maxDecimalesTarifa = count;

                                matrizValores[j, i] = tarifaTramo;
                                j++;
                            }
                        }
                    }
                    i++;
                }

                for (i = 0; i < matrizValores.GetLength(0); i++)
                {
                    for (j = 0; j < matrizValores.GetLength(1); j++)
                    {
                        if (!String.IsNullOrEmpty(matrizValores[i, j]))
                        {
                            double tarifa = double.Parse(matrizValores[i, j]);
                            matrizValores[i, j] = tarifa.ToString("N" + maxDecimalesTarifa) + '€';
                        }
                    }
                }


                //Se buscan los datos de los VA            
                ListaTarifasVAReporte = this.CrearTablaVA(objProductoBE.idProducto, objProductoOfertaBE.idProductoOferta, tarifasConDescuento);

                #endregion

                #region Generacion Plantilla
                const int NUM_BLOQUES_DESTINO = 4;
       

                //Copiamos la plantilla en el fichero temporal.
                File.Copy(rutaPlantilla, ficheroTemporalWord, true);

                //Abrimos el fichero
                objWord.AbrirFichero();
                if (objWord.Abierto)
                {
                    objWord.ActivarModoWord();

                    #region Etiquetas word

                    string textoParametro = string.Empty;
                    if (!tarifasConDescuento)
                    {
                        textoParametro = SimuladorResources.TituloReportTarifas;
                    }
                    else
                    {
                        textoParametro = SimuladorResources.TituloReportPrecios;
                    }

                    //Sustituimos las etiquetas,
                    String saltoLinea = ""; //(ListaTarifasVAReporte.Count > 0) ? "\f" : String.Empty;

                    objEtiquetas.Add("$$TITULOREPORTE", textoParametro);
                    objEtiquetas.Add("$$CODIGOSAP", string.Format(CultureInfo.InvariantCulture, "{0}", objProductoBE.Descripcion));
                    objEtiquetas.Add("$$VALIDEZDESDE", fechaInicial);
                    objEtiquetas.Add("$$VALIDEZHASTA", fechaFinal);
                    objEtiquetas.Add("$$SALTOLINEA", saltoLinea);
                    objEtiquetas.Add("$$NOMBRECLIENTECOMERCIAL", nombreCliente);
                    objWord.BuscaYSustituye(objEtiquetas, true, false);

                    #endregion

                    #region Rellenar las tablas

                    //Rellenamos la tabla de tarifas.
                    objWord.SeleccionarTabla(1);

                    string tablaEnString = "";

                    string[] colsDestino = new string[20];
                    string[] colsDescripcion = new string[20];
                    colsDestino[0] = colsDescripcion[0] = string.Empty;
                    colsDestino[4] = colsDescripcion[4] = string.Empty;
                    colsDestino[8] = colsDescripcion[8] = string.Empty;
                    colsDestino[12] = colsDescripcion[12] = string.Empty;
                    colsDestino[16] = colsDescripcion[16] = string.Empty;

                    String lineaDescuentoDestino = "\nDescuento aplicado sobre tarifa" + '\t';
                    String[] descuentosDestino = new String[3];

                    //Si se trata de descuento por destino, hay que añadir el descuento aplicado sobre tarifa
                    if (objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoDestino) && tarifasConDescuento)
                    {

                        int count = 0;

                        foreach (DestinoBE objDestino in listadoDestinos)
                        {
                            //Obtenemos la configuración del destino con tramos
                            var configDestino = listaConfiguracionDestino.FirstOrDefault(t => t.idDestino.Equals(objDestino.idDestino));

                            //Si el destino tiene configuración y descuento final, mostramos el descuento en la tabla
                            if (configDestino != null && configDestino.DescuentoFinal.HasValue)
                            {
                                descuentosDestino[count] = configDestino.DescuentoFinal.Value.ToString("N2") + "%";
                            }
                            else
                            {
                                descuentosDestino[count] = "0.00%";
                            }

                            count++;
                        }


                        for (int iBloque = 0; iBloque < NUM_BLOQUES_DESTINO; iBloque++)
                        {
                            lineaDescuentoDestino += String.Join("\t", descuentosDestino) + "\t\t";
                        }

                        // Borramos el caracter extra insertado en el for
                        lineaDescuentoDestino = lineaDescuentoDestino.Substring(0, lineaDescuentoDestino.Length - 1);
                    }

                    int destino = 1;
                    foreach (DestinoBE objDestino in listadoDestinos)
                    {
                        //Obtenemos la descripcion
                        var objDescDestino = listaDestinos.FirstOrDefault(d => d.CodDestinoSAP.Equals(objDestino.CodDestinoSAPSinZona));
                        var descDestino = (objDescDestino != null && !String.IsNullOrEmpty(objDescDestino.DescripcionDestino)) ? objDescDestino.DescripcionDestino : "";
                        descDestino = StringUtil.ToTitleCaseIfAllUpper(descDestino);

                        colsDestino[destino] = objDestino.CodDestinoSAPSinZona;
                        colsDescripcion[destino - 1] = descDestino;
                        destino++;
                    }
                    destino++;

                    foreach (DestinoBE objDestino in listadoDestinos)
                    {
                        //Obtenemos la descripcion
                        var objDescDestino = listaDestinos.FirstOrDefault(d => d.CodDestinoSAP.Equals(objDestino.CodDestinoSAPSinZona));
                        var descDestino = (objDescDestino != null && !String.IsNullOrEmpty(objDescDestino.DescripcionDestino)) ? objDescDestino.DescripcionDestino : "";
                        descDestino = StringUtil.ToTitleCaseIfAllUpper(descDestino);

                        colsDestino[destino] = objDestino.CodDestinoSAPSinZona;
                        colsDescripcion[destino - 1] = descDestino;
                        destino++;

                    }
                    destino++;

                    foreach (DestinoBE objDestino in listadoDestinos)
                    {
                        //Obtenemos la descripcion
                        var objDescDestino = listaDestinos.FirstOrDefault(d => d.CodDestinoSAP.Equals(objDestino.CodDestinoSAPSinZona));
                        var descDestino = (objDescDestino != null && !String.IsNullOrEmpty(objDescDestino.DescripcionDestino)) ? objDescDestino.DescripcionDestino : "";
                        descDestino = StringUtil.ToTitleCaseIfAllUpper(descDestino);

                        colsDestino[destino] = objDestino.CodDestinoSAPSinZona;
                        colsDescripcion[destino - 1] = descDestino;
                        destino++;

                    }

                    destino++;

                    foreach (DestinoBE objDestino in listadoDestinos)
                    {
                        //Obtenemos la descripcion
                        var objDescDestino = listaDestinos.FirstOrDefault(d => d.CodDestinoSAP.Equals(objDestino.CodDestinoSAPSinZona));
                        var descDestino = (objDescDestino != null && !String.IsNullOrEmpty(objDescDestino.DescripcionDestino)) ? objDescDestino.DescripcionDestino : "";
                        descDestino = StringUtil.ToTitleCaseIfAllUpper(descDestino);

                        colsDestino[destino] = objDestino.CodDestinoSAPSinZona;
                        colsDescripcion[destino - 1] = descDestino;
                        destino++;

                    }
                    //destino++;

                    //foreach (DestinoBE objDestino in listadoDestinos)
                    //{
                    //    //Obtenemos la descripcion
                    //    var objDescDestino = listaDestinos.FirstOrDefault(d => d.CodDestinoSAP.Equals(objDestino.CodDestinoSAPSinZona));
                    //    var descDestino = (objDescDestino != null && !String.IsNullOrEmpty(objDescDestino.DescripcionDestino)) ? objDescDestino.DescripcionDestino : "";
                    //    descDestino = StringUtil.ToTitleCaseIfAllUpper(descDestino);

                    //    colsDestino[destino] = objDestino.CodDestinoSAPSinZona;
                    //    colsDescripcion[destino - 1] = descDestino;
                    //    destino++;
                            
                    //}

                    String tablaDescripcion = "\t";

                    foreach (string s in colsDestino)
                    {
                        tablaEnString += s + '\t';
                    }

                    //Es Informe de Precios DD
                    var esInformePreciosDD = (objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoDestino) && tarifasConDescuento);
                    
                    if (!esInformePreciosDD)
                    {
                        tablaDescripcion = "Peso\t";
                    }
                    
                    foreach(string s in colsDescripcion)
                    {
                        tablaDescripcion += s + "\t";
                    }

                    // Borramos el caracter extra insertado en el for
                    //tablaEnString = tablaEnString.Substring(0, tablaEnString.Length - 1);
                    //tablaEnString +=  "\n" + tablaDescripcion.Substring(0, tablaDescripcion.Length - 1);
                    tablaEnString = tablaEnString.TrimEnd('\t');
                    tablaEnString += "\n" + tablaDescripcion.TrimEnd('\t');

                    //Si se trata de descuento por destino, hay que añadir el descuento aplicado sobre tarifa
                    if (objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoDestino) && tarifasConDescuento)
                    {
                        // Borramos los caracter extra insertados
                        lineaDescuentoDestino = lineaDescuentoDestino.Substring(0, lineaDescuentoDestino.Length - 1);
                        lineaDescuentoDestino += '\n';
                        lineaDescuentoDestino += "Peso\tPRECIO FINAL";

                        tablaEnString += lineaDescuentoDestino;
                    }

                    int iteraciones = 1 + (numTramos / NUM_BLOQUES_DESTINO);

                    for (int k = 0; k < iteraciones; k++)
                    {
                        tablaEnString += '\n';

                        colsDestino = new string[16];

                        int indiceTablaA = k;
                        int indiceTablaB = k + iteraciones;
                        int indiceTablaC = k + (2 * iteraciones);
                        int indiceTablaD = k + (3 * iteraciones);
                        //int indiceTablaE = k + (4 * iteraciones);

                        if (indiceTablaA < numTramos)
                        {
                            colsDestino[0] = auxTramo[indiceTablaA].Descripcion.Replace(" GRS", "g").Replace("GR", "g");
                            colsDestino[1] = matrizValores[indiceTablaA, 0];
                            colsDestino[2] = matrizValores[indiceTablaA, 1];
                            colsDestino[3] = matrizValores[indiceTablaA, 2];
                        }
                        else
                        {
                            colsDestino[0] = string.Empty;
                            colsDestino[1] = string.Empty;
                            colsDestino[2] = string.Empty;
                            colsDestino[3] = string.Empty;
                        }

                        if (indiceTablaB < numTramos)
                        {
                            colsDestino[4] = auxTramo[indiceTablaB].Descripcion.Replace(" GRS", "g").Replace("GR", "g");
                            colsDestino[5] = matrizValores[indiceTablaB, 0];
                            colsDestino[6] = matrizValores[indiceTablaB, 1];
                            colsDestino[7] = matrizValores[indiceTablaB, 2];
                        }
                        else
                        {
                            colsDestino[4] = string.Empty;
                            colsDestino[5] = string.Empty;
                            colsDestino[6] = string.Empty;
                            colsDestino[7] = string.Empty;
                        }

                        if (indiceTablaC < numTramos)
                        {
                            colsDestino[8] = auxTramo[indiceTablaC].Descripcion.Replace(" GRS", "g").Replace("GR", "g");
                            colsDestino[9] = matrizValores[indiceTablaC, 0];
                            colsDestino[10] = matrizValores[indiceTablaC, 1];
                            colsDestino[11] = matrizValores[indiceTablaC, 2];
                        }
                        else
                        {
                            colsDestino[8] = string.Empty;
                            colsDestino[9] = string.Empty;
                            colsDestino[10] = string.Empty;
                            colsDestino[11] = string.Empty;
                        }

                        if (indiceTablaD < numTramos)
                        {
                            colsDestino[12] = auxTramo[indiceTablaD].Descripcion.Replace(" GRS", "g").Replace("GR", "g");
                            colsDestino[13] = matrizValores[indiceTablaD, 0];
                            colsDestino[14] = matrizValores[indiceTablaD, 1];
                            colsDestino[15] = matrizValores[indiceTablaD, 2];
                        }
                        else
                        {
                            colsDestino[12] = string.Empty;
                            colsDestino[13] = string.Empty;
                            colsDestino[14] = string.Empty;
                            colsDestino[15] = string.Empty;
                        }

                        //if (indiceTablaE < numTramos)
                        //{
                        //    colsDestino[16] = auxTramo[indiceTablaE].Descripcion.Replace(" GRS", "g").Replace("GR", "g");
                        //    colsDestino[17] = matrizValores[indiceTablaE, 0];
                        //    colsDestino[18] = matrizValores[indiceTablaE, 1];
                        //    colsDestino[19] = matrizValores[indiceTablaE, 2];
                        //}
                        //else
                        //{
                        //    colsDestino[16] = string.Empty;
                        //    colsDestino[17] = string.Empty;
                        //    colsDestino[18] = string.Empty;
                        //    colsDestino[19] = string.Empty;
                        //}

                        foreach (string s in colsDestino)
                        {
                            tablaEnString += s + '\t';
                        }

                        // Borramos el caracter extra insertado en el for
                        tablaEnString = tablaEnString.Substring(0, tablaEnString.Length - 1);
                    }

                    objWord.EscribirTablaDeString(tablaEnString, false, maxDecimalesTarifa);

                    objWord.PonerTablaCabeceraRepetidaPublicorreo();

                    if (objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoDestino))
                    {
                        objWord.DarFormatoTablaDescuentoDestino(esInformeTarifas: !tarifasConDescuento ,esPublicorreo:true);
                    }
                    else if (objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoTramo) ||
                            objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoPrecioCierto))
                    {
                         objWord.DarFormatoTablaPrecioCierto(esPublicorreo:true);
                    }

                    objWord.EscribirTextoAlFinalDelDocumento("$$TABLAVA");

                    //Rellenamos la tabla de VA.
                    if (ListaTarifasVAReporte.Count > 0)
                    {
                        //insertamos la plantilla de la tabla de VA
                        objWord.InsertarDocumentoSustituyendoTexto(rutaPlantillaVA, "$$TABLAVA", (int)FormasSustitucionTexto.CopiarPegarConFormato);

                        objWord.SeleccionarTabla(2);
                        string[] colsVA = new string[2];
                        //colsVA[0] = "LISTADO DE VALORES AÑADIDOS";
                        //colsVA[1] = string.Empty;
                        //objWord.AgregarTituloSinFormatoATabla(colsVA);
                        foreach (ReporteVABE objVA in ListaTarifasVAReporte)
                        {
                            colsVA[0] = objVA.Nombre;
                            colsVA[1] = objVA.Descripcion;
                            objWord.AgregarFilaSinFormatoATabla(colsVA);
                        }
                        //objWord.EliminarFilaDeTabla(2,2);
                        objWord.PonerTablaCabeceraRepetida();
                        objWord.DarFormatoTablaVA();                       
                    }
                    else
                    {
                        objWord.BuscaYSustituye("$$TABLAVA", "$$OBSERVACIONES", true, false);
                    }

                    #endregion

                    #region insertar pie de informe según linea de producto

                    objWord.InsertarDocumentoSustituyendoTexto(rutaLineaProducto, "$$OBSERVACIONES", (int)FormasSustitucionTexto.CopiarPegarConFormato);
                    saltoLinea = (ListaTarifasVAReporte.Count > 0) ? "\f" : "";
                    objWord.BuscaYSustituye("$$SALTOLINEA", saltoLinea, true, false);

                    //objWord.InsertarDocumentoSustituyendoTexto(rutaLineaProducto, "$$OBSERVACIONES", (int)FormasSustitucionTexto.CopiarPegarConFormato);

                    InformacionDestinosBL objInfoDestinosBL = new InformacionDestinosBL();
                    Collection<InformacionDestinosBE> objListaDestinos = objInfoDestinosBL.ObtenerListadoInformacionDestinos(objProductoBE.CodProducto);
                    StringBuilder sb = new StringBuilder();
                    if (objListaDestinos != null)
                    {
                        int insertados = 1;
                        foreach (DestinoBE objDestino in listadoDestinos)
                        {
                            InformacionDestinosBE objDescripcion = objListaDestinos.FirstOrDefault(x => x.CodDestinoSAP.Equals(objDestino.CodDestinoSAP));

                            if ((objDescripcion != null) && (!string.IsNullOrWhiteSpace(objDescripcion.DescripcionDestino)))
                            {
                                if (insertados < numDestinos)
                                {
                                    sb.Append(string.Format(CultureInfo.InvariantCulture, "{0}: {1}, ", objDescripcion.CodDestinoSAP, objDescripcion.DescripcionDestino));
                                }
                                else
                                {
                                    sb.Append(string.Format(CultureInfo.InvariantCulture, "{0}: {1}. ", objDescripcion.CodDestinoSAP, objDescripcion.DescripcionDestino));
                                }
                                insertados++;
                            }
                        }

                        //En caso de que la lista de destinos del producto y de la DB difieran.
                        if (sb.ToString() != string.Empty) sb.Replace(',', '.', sb.Length - 2, 1);
                    }
                   
                    objWord.BuscaYSustituye(objEtiquetas, true, false);

                    if (informeMultiple && !esUltimoInforme)                   
                        objWord.InsertarCambioSeccionAlFinalDocumento();

                    #endregion

                    #region Generar WORD

                    //Ya tenemos generado el documento, ahora lo guardamos en PDF.                    
                    ficheroDestino = string.Empty;

                    if (tarifasConDescuento)
                    {
                        ficheroDestino = Path.Combine(System.IO.Path.GetTempPath(), string.Format(CultureInfo.InvariantCulture, "{0}_{7}_{1}{2}{3}{4}{5}{6}.docx", "InformePrecios", DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(),
                            DateTime.Now.Day.ToString(), DateTime.Now.Hour.ToString(), DateTime.Now.Minute.ToString(), DateTime.Now.Second.ToString(), objProductoBE.CodProducto));
                    }
                    else
                    {
                        ficheroDestino = Path.Combine(System.IO.Path.GetTempPath(), string.Format(CultureInfo.InvariantCulture, "{0}_{7}_{1}{2}{3}{4}{5}{6}.docx", "InformeTarifas", DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(),
                            DateTime.Now.Day.ToString(), DateTime.Now.Hour.ToString(), DateTime.Now.Minute.ToString(), DateTime.Now.Second.ToString(), objProductoBE.CodProducto));

                    }
                    //Guardamos el fichero
                    if (informeMultiple)
                    {
                        objWord.GuardarComo(ficheroTemporalPdf);
                    }
                    else
                    {
                        objWord.GuardarComo(ficheroDestino);
                    }

                    #endregion
                }

                #endregion

            }
            catch (Exception ex)
            {
                RegistrarAccionesSimulador.GuardarExcepcion(ex);
            }
            finally
            {
                if (objWord.Abierto)
                {
                    objWord.CerrarWord();
                }

                if (!informeMultiple)
                {
                    //Una vez creado el fichero en temporal se abre.
                    if (System.IO.File.Exists(ficheroDestino))
                    {
                        System.Diagnostics.Process proc = new System.Diagnostics.Process();
                        proc.EnableRaisingEvents = false;
                        proc.StartInfo.FileName = ficheroDestino;
                        proc.Start();
                    }
                }
            }
        }

        /// <summary>
        /// genera un informe de tarifas/precios en formato pdf con modo de publicorreo
        /// </summary>
        /// <param name="objProductoBE"></param>
        /// <param name="objProductoOfertaBE"></param>
        /// <param name="fechaInicial"></param>
        /// <param name="fechaFinal"></param>
        /// <param name="nombreCliente"></param>
        /// <param name="tarifasConDescuento"></param>
        /// <param name="esUltimoInforme">indica si es el último producto de un informe múltimple</param>
        /// <returns>Devuelve la ruta del fichero dónde se guarda el informe.</returns>
        private void GenerarInformeEstandarPubliCorreoExcel(ProductoBE objProductoBE, ProductoOfertaBE objProductoOfertaBE, string fechaInicial, string fechaFinal, string nombreCliente, bool tarifasConDescuento, bool informeMultiple, string ficheroTemporalExcel)
        {
            #region Variables

            //En caso de ser un producto con Destinos se rellena esta matriz de informacion
            object[,] matrizValores = null;
            bool esTipoPrecioCierto = false;

            //Listado de valores que se ingresan en la tabla de valores añadidos si corresponde
            Collection<ReporteVABE> ListaTarifasVAReporte = new Collection<ReporteVABE>();

            //Lista de etiquetas con su correspondiente valor que se sustituye en el documento word
            Dictionary<string, string> objEtiquetas = new Dictionary<string, string>();

            //Contiene la ruta de la plantilla que se usa para generar el reporte de tarifas
            string rutaPlantilla = string.Format(Utils.GetValorFromAppConfig(AppSettingsEnum.RutaPlantillaInformeTarifas), AppDomain.CurrentDomain.BaseDirectory);
            rutaPlantilla = rutaPlantilla.Substring(0, rutaPlantilla.Length - 4) + "xlsx";

            string rutaLineaProducto = string.Format(Utils.GetValorFromAppConfig(AppSettingsEnum.RutaPlantillaLineaProducto), AppDomain.CurrentDomain.BaseDirectory, objProductoBE.PlantillaInformeTarifasPrecios);
            rutaLineaProducto = rutaLineaProducto.Substring(0, rutaLineaProducto.Length - 4) + "xlsx";

            string rutaPlantillaVA = string.Format(Utils.GetValorFromAppConfig(AppSettingsEnum.RutaPlantillaInformeTarifasVA), AppDomain.CurrentDomain.BaseDirectory);

            // Si no es un informe múltiple o no tiene nombre, poner nombre por defecto
            if (ficheroTemporalExcel.Equals(string.Empty))// || !informeMultiple)
            {
                if (tarifasConDescuento)
                {
                    ficheroTemporalExcel = Path.Combine(System.IO.Path.GetTempPath(), string.Format(CultureInfo.InvariantCulture, "{0}_{7}_{1}{2}{3}{4}{5}{6}.xlsx", "InformePrecios", DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(),
                        DateTime.Now.Day.ToString(), DateTime.Now.Hour.ToString(), DateTime.Now.Minute.ToString(), DateTime.Now.Second.ToString(), objProductoBE.CodProducto));
                }
                else
                {
                    ficheroTemporalExcel = Path.Combine(System.IO.Path.GetTempPath(), string.Format(CultureInfo.InvariantCulture, "{0}_{7}_{1}{2}{3}{4}{5}{6}.xlsx", "InformeTarifas", DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(),
                       DateTime.Now.Day.ToString(), DateTime.Now.Hour.ToString(), DateTime.Now.Minute.ToString(), DateTime.Now.Second.ToString(), objProductoBE.CodProducto));

                }
            }

            //instancia del objeto word con el que vamos  trabajar para generar el informe de tarifas
            ManagerExcel objExcel = new ManagerExcel(ficheroTemporalExcel, false);

            //Lista de tramos que usamos para, al generar el informe, mostrar la cabecera de los tramos.
            List<TramoBE> auxTramo = new List<TramoBE>();

            Collection<TramoBE> listaTramosEliminar = new Collection<TramoBE>();

            #endregion

            try
            {
                #region Obtencion Datos

                // Guardamos si es precio cierto o tipo descuento
                if (objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoPrecioCierto) || objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorPrecioCiertoGrupoTramo))
                {
                    esTipoPrecioCierto = true;
                }
                else
                {
                    esTipoPrecioCierto = false;
                }

                //Obtenemos los datos que ha ingresado el usuario.
                ConfiguracionProductosBL configProducto = new ConfiguracionProductosBL();
                Collection<ConfiguracionDestinoOfertaBE> listaConfiguracionDestino = configProducto.ObtenerConfiguracionDestinoOferta(objProductoOfertaBE.idProductoOferta);
                Collection<ConfiguracionTramoOfertaBE> listaConfiguracionTramo = configProducto.ObtenerConfiguracionTramoOferta(objProductoOfertaBE.idProductoOferta);
                Collection<InformacionDestinosBE> listaDestinos = new InformacionDestinosBL().ObtenerListadoInformacionDestinos(objProductoBE.CodProducto);
                List<TramoInformeBE> listaTramosInfTarifas = new TramoInformeBL().ObtenerTramosInformeProducto(objProductoBE.CodProducto);

                var listadoDestinos = objProductoBE.Destinos.Where(x => x.Tramos.Count > 0).OrderBy(t => t.Orden).ToList();

                if (objProductoBE.Internacional)
                {
                    listadoDestinos = (from d in listadoDestinos
                                       join c in listaConfiguracionDestino on d.idDestino equals c.idDestino
                                       where c.Distribucion.HasValue && c.Distribucion.Value > 0
                                       select d).ToList();

                }

                int numDestinos = listadoDestinos.Count();

                //Producto con destinos
                int numTramos = 0;
                //obtenemos el número máximo de tramos de entre los destinos. Ademas nos guardamos los tramos para la cabecera de fila de la tabla del informe. 
                foreach (DestinoBE destino in listadoDestinos)
                {
                    if (numTramos < destino.Tramos.Count)
                    {
                        numTramos = destino.Tramos.Count;
                        //auxTramo = destino.Tramos;
                        TramoBE[] copiaTramos = new TramoBE[numTramos];
                        destino.Tramos.CopyTo(copiaTramos, 0);
                        auxTramo = copiaTramos.ToList();
                    }
                }

                decimal limiteTramos = 250;

                foreach (TramoBE item in auxTramo.OrderBy(x => x.CodTramo))
                {
                    if (item.CodTramoDecimal > limiteTramos)
                    {
                        //[FIX][MMUNOZ] Quieren que se muestren todos los valores de los tramos, tengan distribución o no
                        Collection<ConfiguracionTramoOfertaBE> auxRellenos = listaConfiguracionTramo.Where(x => (x.CodTramo.Equals(item.CodTramo)) && (x.Distribucion.HasValue) && (x.Distribucion.Value > 0)).ToList<ConfiguracionTramoOfertaBE>().ToCollection<ConfiguracionTramoOfertaBE>();
                    
                        if ((auxRellenos == null) || (auxRellenos.Count.Equals(0)))
                        {
                            listaTramosEliminar.Add(item);
                        }
                        else
                        {
                            listaTramosEliminar.Clear();
                        }
                    }
                }

                numTramos = auxTramo.Count;
                //Eliminamos los tramos que no deben insertarse
                foreach (TramoBE item in listaTramosEliminar)
                {
                    numTramos--;
                }

                matrizValores = new object[numTramos, numDestinos];
                int i = 0;
                int j = 0;
                decimal maxDecimalesTarifa = 2;

                foreach (DestinoBE destino in listadoDestinos)
                {
                    j = 0;
                    foreach (TramoBE tramo in destino.Tramos)
                    {
                        TramoBE auxEliminar = listaTramosEliminar.FirstOrDefault(x => x.CodTramo.Equals(tramo.CodTramo));
                        if (auxEliminar == null)
                        {
                            TramoBE aux = auxTramo.FirstOrDefault(x => x.CodTramo.Equals(tramo.CodTramo));
                            if (aux != null)
                            {
                                object tarifaTramo = Math.Round(tramo.Tarifa, 5);

                                if (esTipoPrecioCierto)
                                {
                                    tarifaTramo = Math.Round(tramo.Tarifa, 5);
                                }

                                if (tarifasConDescuento)
                                {
                                    double auxTarifa = 0;

                                    //Según el tipo de modalidad se calcula la tarifa de una forma u otra.
                                    if (objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoDestino))
                                    {
                                        //Se obtiene la configuración del destino
                                        ConfiguracionDestinoOfertaBE configDestino = listaConfiguracionDestino.FirstOrDefault(x => x.idDestino.Equals(tramo.idDestino.Value));

                                        if (configDestino != null)
                                        {
                                            //Si se quiere mostrar la tarifa con los descuentos
                                            if (configDestino.DescuentoFinal.HasValue && double.TryParse(configDestino.DescuentoFinal.Value.ToString(), out auxTarifa))
                                            {
                                                auxTarifa = tramo.Tarifa - (tramo.Tarifa * auxTarifa / 100);
                                                tarifaTramo = Math.Round(auxTarifa, 5);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        //Se obtiene la configuración del tramo
                                        ConfiguracionTramoOfertaBE configTramo = listaConfiguracionTramo.FirstOrDefault(x => x.idTramo.Equals(tramo.idTramo));

                                        if (configTramo != null)
                                        {
                                            //Si se quiere mostrar la tarifa con los descuentos
                                            if ((objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoTramo)) &&
                                                    configTramo.DescuentoFinal.HasValue && double.TryParse(configTramo.DescuentoFinal.Value.ToString(), out auxTarifa))
                                            {
                                                auxTarifa = tramo.Tarifa - (tramo.Tarifa * auxTarifa / 100);

                                                tarifaTramo = Math.Round(auxTarifa, 5);
                                            }
                                            else if ((objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoPrecioCierto)) &&
                                                    configTramo.PrecioCierto.HasValue && double.TryParse(configTramo.PrecioCierto.Value.ToString(), out auxTarifa))
                                            {
                                                if (!auxTarifa.Equals(0))
                                                {
                                                    tarifaTramo = Math.Round(auxTarifa, 5);
                                                }
                                            }
                                        }
                                    }
                                }
                                
                                decimal argument = (decimal)(double)tarifaTramo;
                                int count = BitConverter.GetBytes(decimal.GetBits(argument)[3])[2];

                                if (count >= maxDecimalesTarifa)
                                    maxDecimalesTarifa = count;

                                matrizValores[j, i] = tarifaTramo;
                                j++;
                            }
                        }
                    }
                    i++;
                }

                //Se buscan los datos de los VA            
                ListaTarifasVAReporte = this.CrearTablaVA(objProductoBE.idProducto, objProductoOfertaBE.idProductoOferta, tarifasConDescuento);

                #endregion

                #region Generacion Plantilla

                if (!File.Exists(ficheroTemporalExcel))
                {
                    //Copiamos la plantilla en el fichero temporal.
                    File.Copy(rutaPlantilla, ficheroTemporalExcel, true);
                }

                //Abrimos el fichero
                try
                {
                    objExcel.AbrirFichero();
                }
                catch { }
                if (objExcel.Abierto)
                {
                    objExcel.SeleccionarPrimeraHojaLibreOCrearNuevaYRenombrar(objProductoBE.CodAnexoSAP + " - " + objProductoBE.CodProducto + " - " + GenerarAbreviaturaModeloNegociacion(objProductoOfertaBE.CodModalidadNegociacion.Trim()));

                    //if (informeMultiple && esNecesarioInsertarPlantilla) { objExcel.InsertarDocumentoAlFinal(rutaPlantilla, 1); }

                    #region "Borrar una vez validado el método"

                    #endregion

                    #region Etiquetas excel

                    string textoParametro = string.Empty;
                    if (!tarifasConDescuento)
                    {
                        textoParametro = SimuladorResources.TituloReportTarifas;
                    }
                    else
                    {
                        textoParametro = SimuladorResources.TituloReportPrecios;
                    }

                    #endregion

                    #region Rellenar las tablas

                    objExcel.IniciarDibujarTabla(true, true);

                    string[] colsDestinoHeader = new string[16];
                    string[] colsDescripcion = new string[16];
                    colsDestinoHeader[0] = colsDescripcion[0] = string.Empty;
                    colsDestinoHeader[4] = colsDescripcion[4] = string.Empty;
                    colsDestinoHeader[8] = colsDescripcion[8] = string.Empty;
                    colsDestinoHeader[12] = colsDescripcion[12] = string.Empty;
                    //colsDestinoHeader[16] = colsDescripcion[16] = string.Empty;


                    //Es Informe de Precios DD
                    var esInformePreciosDD = (objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoDestino) && tarifasConDescuento);

                    if (!esInformePreciosDD)
                    {
                        colsDescripcion[0] = "Peso";
                    }

                    int destino = 1;
                    foreach (DestinoBE objDestino in listadoDestinos)
                    {
                        //Obtenemos la descripcion
                        var objDescDestino = listaDestinos.FirstOrDefault(d => d.CodDestinoSAP.Equals(objDestino.CodDestinoSAPSinZona));
                        var descDestino = (objDescDestino != null && !String.IsNullOrEmpty(objDescDestino.DescripcionDestino)) ? objDescDestino.DescripcionDestino : "";
                        descDestino = StringUtil.ToTitleCaseIfAllUpper(descDestino);

                        colsDestinoHeader[destino] = objDestino.CodDestinoSAPSinZona;
                        colsDescripcion[destino] = descDestino;
                        destino++;
                    }
                    destino++;

                    foreach (DestinoBE objDestino in listadoDestinos)
                    {
                        //Obtenemos la descripcion
                        var objDescDestino = listaDestinos.FirstOrDefault(d => d.CodDestinoSAP.Equals(objDestino.CodDestinoSAPSinZona));
                        var descDestino = (objDescDestino != null && !String.IsNullOrEmpty(objDescDestino.DescripcionDestino)) ? objDescDestino.DescripcionDestino : "";
                        descDestino = StringUtil.ToTitleCaseIfAllUpper(descDestino);

                        colsDestinoHeader[destino] = objDestino.CodDestinoSAPSinZona;
                        colsDescripcion[destino] = descDestino;
                        destino++;
                    }
                    destino++;

                    foreach (DestinoBE objDestino in listadoDestinos)
                    {
                        //Obtenemos la descripcion
                        var objDescDestino = listaDestinos.FirstOrDefault(d => d.CodDestinoSAP.Equals(objDestino.CodDestinoSAPSinZona));
                        var descDestino = (objDescDestino != null && !String.IsNullOrEmpty(objDescDestino.DescripcionDestino)) ? objDescDestino.DescripcionDestino : "";
                        descDestino = StringUtil.ToTitleCaseIfAllUpper(descDestino);

                        colsDestinoHeader[destino] = objDestino.CodDestinoSAPSinZona;
                        colsDescripcion[destino] = descDestino;
                        destino++;
                    }

                    destino++;

                    foreach (DestinoBE objDestino in listadoDestinos)
                    {
                        //Obtenemos la descripcion
                        var objDescDestino = listaDestinos.FirstOrDefault(d => d.CodDestinoSAP.Equals(objDestino.CodDestinoSAPSinZona));
                        var descDestino = (objDescDestino != null && !String.IsNullOrEmpty(objDescDestino.DescripcionDestino)) ? objDescDestino.DescripcionDestino : "";
                        descDestino = StringUtil.ToTitleCaseIfAllUpper(descDestino);

                        colsDestinoHeader[destino] = objDestino.CodDestinoSAPSinZona;
                        colsDescripcion[destino] = descDestino;
                        destino++;
                    }
                    //destino++;

                    //foreach (DestinoBE objDestino in listadoDestinos)
                    //{
                    //    //Obtenemos la descripcion
                    //    var objDescDestino = listaDestinos.FirstOrDefault(d => d.CodDestinoSAP.Equals(objDestino.CodDestinoSAPSinZona));
                    //    var descDestino = (objDescDestino != null && !String.IsNullOrEmpty(objDescDestino.DescripcionDestino)) ? objDescDestino.DescripcionDestino : "";
                    //    descDestino = StringUtil.ToTitleCaseIfAllUpper(descDestino);

                    //    colsDestinoHeader[destino] = objDestino.CodDestinoSAPSinZona;
                    //    colsDescripcion[destino] = descDestino;
                    //    destino++;
                    //}

                    objExcel.AgregarFila(colsDestinoHeader);
                    objExcel.AgregarFila(colsDescripcion, true);

                    //Si se trata de descuento por destino, hay que añadir el descuento aplicado sobre tarifa
                    if (objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoDestino) && tarifasConDescuento)
                    {
                        int contadorCols = 1;                        
                        colsDestinoHeader = new string[16];
                        colsDestinoHeader[0] += "Descuento aplicado sobre tarifa";

                        for (int z = 0; z < 4; z++)
                        {
                            foreach (DestinoBE objDestino in listadoDestinos)
                            {
                                //Obtenemos la configuración del destino con tramos
                                var configDestino = listaConfiguracionDestino.FirstOrDefault(t => t.idDestino.Equals(objDestino.idDestino));

                                //Si el destino tiene configuración y descuento final, mostramos el descuento en la tabla
                                if (configDestino != null && configDestino.DescuentoFinal.HasValue)
                                {
                                    colsDestinoHeader[contadorCols] = configDestino.DescuentoFinal.Value.ToString("N2") + "%";
                                }
                                else
                                {
                                    colsDestinoHeader[contadorCols] = "0,00%";
                                }

                                contadorCols++;
                            }

                            if(contadorCols < colsDestinoHeader.Length - 1)
                                colsDestinoHeader[contadorCols] = "";
                            
                            contadorCols++;
                        }
                        
                        objExcel.AgregarFila(colsDestinoHeader);
                        objExcel.AgregarFila(new String[] { "Peso", "PRECIO FINAL" });
                    }
                    


                    int iteraciones = 1 + (numTramos / 4);

                    for (int k = 0; k < iteraciones; k++)
                    {
                        object[] colsDestino = new object[16];

                        int indiceTablaA = k;
                        int indiceTablaB = k + iteraciones;
                        int indiceTablaC = k + (2 * iteraciones);
                        int indiceTablaD = k + (3 * iteraciones);
                        //int indiceTablaE = k + (4 * iteraciones);

                        if (indiceTablaA < numTramos)
                        {
                            colsDestino[0] = auxTramo[indiceTablaA].Descripcion.Replace(" GRS", "g").Replace("GR", "g");
                            colsDestino[1] = matrizValores[indiceTablaA, 0];
                            colsDestino[2] = matrizValores[indiceTablaA, 1];
                            colsDestino[3] = matrizValores[indiceTablaA, 2];
                        }
                        else
                        {
                            colsDestino[0] = string.Empty;
                            colsDestino[1] = string.Empty;
                            colsDestino[2] = string.Empty;
                            colsDestino[3] = string.Empty;
                        }

                        if (indiceTablaB < numTramos)
                        {
                            colsDestino[4] = auxTramo[indiceTablaB].Descripcion.Replace(" GRS", "g").Replace("GR", "g");
                            colsDestino[5] = matrizValores[indiceTablaB, 0];
                            colsDestino[6] = matrizValores[indiceTablaB, 1];
                            colsDestino[7] = matrizValores[indiceTablaB, 2];
                        }
                        else
                        {
                            colsDestino[4] = string.Empty;
                            colsDestino[5] = string.Empty;
                            colsDestino[6] = string.Empty;
                            colsDestino[7] = string.Empty;
                        }

                        if (indiceTablaC < numTramos)
                        {
                            colsDestino[8] = auxTramo[indiceTablaC].Descripcion.Replace(" GRS", "g").Replace("GR", "g");
                            colsDestino[9] = matrizValores[indiceTablaC, 0];
                            colsDestino[10] = matrizValores[indiceTablaC, 1];
                            colsDestino[11] = matrizValores[indiceTablaC, 2];
                        }
                        else
                        {
                            colsDestino[8] = string.Empty;
                            colsDestino[9] = string.Empty;
                            colsDestino[10] = string.Empty;
                            colsDestino[11] = string.Empty;
                        }

                        if (indiceTablaD < numTramos)
                        {
                            colsDestino[12] = auxTramo[indiceTablaD].Descripcion.Replace(" GRS", "g").Replace("GR", "g");
                            colsDestino[13] = matrizValores[indiceTablaD, 0];
                            colsDestino[14] = matrizValores[indiceTablaD, 1];
                            colsDestino[15] = matrizValores[indiceTablaD, 2];
                        }
                        else
                        {
                            colsDestino[12] = string.Empty;
                            colsDestino[13] = string.Empty;
                            colsDestino[14] = string.Empty;
                            colsDestino[15] = string.Empty;
                        }

                        //if (indiceTablaE < numTramos)
                        //{
                        //    colsDestino[16] = auxTramo[indiceTablaE].Descripcion.Replace(" GRS", "g").Replace("GR", "g");
                        //    colsDestino[17] = matrizValores[indiceTablaE, 0];
                        //    colsDestino[18] = matrizValores[indiceTablaE, 1];
                        //    colsDestino[19] = matrizValores[indiceTablaE, 2];
                        //}
                        //else
                        //{
                        //    colsDestino[16] = string.Empty;
                        //    colsDestino[17] = string.Empty;
                        //    colsDestino[18] = string.Empty;
                        //    colsDestino[19] = string.Empty;
                        //}

                        objExcel.AgregarFila(colsDestino);
                    }

                    objExcel.PonerColumnaEnNegrita(4);
                    objExcel.PonerColumnaEnNegrita(8);
                    objExcel.PonerColumnaEnNegrita(12);
                    //objExcel.PonerColumnaEnNegrita(16);

                    objExcel.TerminarDibujarTabla(true, false,  (int?) maxDecimalesTarifa);

                    int numColumnasDestinos = colsDestinoHeader.Count() - 1; 

                    //Modificamos el formato de la tabla para la modalidad descuento por destino
                    if (objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoDestino))
                    {   
                        objExcel.DarFormatoTablaDescuentoDestino(numColumnasDestinos, esInformeTarifas: !tarifasConDescuento, esPublicorreo: true);
                    }
                    else if (objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoPrecioCierto) ||
                            objProductoOfertaBE.CodModalidadNegociacion.Equals(SimuladorResources.ValorDescuentoTramo))
                    {
                        objExcel.DarFormatoTablaPrecioCierto(numColumnasDestinos, numTramos, true);
                    }

                    //Rellenamos la tabla de VA.
                    if (ListaTarifasVAReporte.Count > 0)
                    {
                        objExcel.IniciarDibujarTabla(true);
                        string[] colsVATitulo = new string[2];
                        colsVATitulo[0] = "LISTADO DE VALORES AÑADIDOS";
                        colsVATitulo[1] = string.Empty;
                        objExcel.AgregarFila(colsVATitulo, ajustarTextoColumna: false);
                        foreach (ReporteVABE objVA in ListaTarifasVAReporte)
                        {
                            string[] colsVA = new string[2];
                            colsVA[0] = objVA.Nombre;
                            colsVA[1] = objVA.Descripcion;
                            objExcel.AgregarFila(colsVA);
                        }
                        objExcel.TerminarDibujarTabla(false, false, null, true);
                    }
                    else
                    {
                        objExcel.BuscaYSustituye("$$TABLAVA", "$$OBSERVACIONES", true, false);
                    }

                    #endregion

                    #region insertar pie de informe según linea de producto

                    objExcel.InsertarDocumentoAlFinal(rutaLineaProducto);


                    InformacionDestinosBL objInfoDestinosBL = new InformacionDestinosBL();
                    Collection<InformacionDestinosBE> objListaDestinos = objInfoDestinosBL.ObtenerListadoInformacionDestinos(objProductoBE.CodProducto);
                    StringBuilder sb = new StringBuilder();
                    if (objListaDestinos != null)
                    {
                        int insertados = 1;
                        foreach (DestinoBE objDestino in listadoDestinos)
                        {
                            InformacionDestinosBE objDescripcion = objListaDestinos.FirstOrDefault(x => x.CodDestinoSAP.Equals(objDestino.CodDestinoSAP));

                            if ((objDescripcion != null) && (!string.IsNullOrWhiteSpace(objDescripcion.DescripcionDestino)))
                            {
                                if (insertados < numDestinos)
                                {
                                    sb.Append(string.Format(CultureInfo.InvariantCulture, "{0}: {1}, ", objDescripcion.CodDestinoSAP, objDescripcion.DescripcionDestino));
                                }
                                else
                                {
                                    sb.Append(string.Format(CultureInfo.InvariantCulture, "{0}: {1}. ", objDescripcion.CodDestinoSAP, objDescripcion.DescripcionDestino));
                                }
                                insertados++;
                            }
                        }

                        //En caso de que la lista de destinos del producto y de la DB difieran.
                        if (sb.ToString() != string.Empty) sb.Replace(',', '.', sb.Length - 2, 1);
                    }
                  
                    objExcel.BuscaYSustituye(objEtiquetas, true, false);

                    #endregion

                    #region Sustitución de tags

                    // Al contrario que en word, hacemos esto lo último para no tener que insertar la tabla entre medias

                    //Sustituimos las etiquetas,
                    objEtiquetas.Add("$$TITULOREPORTE", textoParametro);
                    objEtiquetas.Add("$$CODIGOSAP", string.Format(CultureInfo.InvariantCulture, "{0}", objProductoBE.Descripcion));
                    objEtiquetas.Add("$$VALIDEZDESDE", fechaInicial);
                    objEtiquetas.Add("$$VALIDEZHASTA", fechaFinal);
                    objEtiquetas.Add("$$NOMBRECLIENTECOMERCIAL", nombreCliente);
                    objExcel.BuscaYSustituye(objEtiquetas, true, false);

                    #endregion
                }

                #endregion

                #region Abrir el fichero

                if (objExcel.Abierto)
                {
                    objExcel.GuardarLibro();
                    objExcel.CerrarExcel();
                }

                if (!informeMultiple)
                {
                    ManagerExcel.AbrirExcelStandalone(ficheroTemporalExcel);
                }

                #endregion
            }
            catch (Exception ex)
            {
                RegistrarAccionesSimulador.GuardarExcepcion(ex);
            }
            finally
            {
                if (objExcel.Abierto)
                {
                    objExcel.CerrarExcel();
                }
            }
        }

        #endregion

        #region GenerarFichasCorreosExpress

        /// <summary>
        /// Genera el listado de fichas de CorreosExpres en la ruta seleccionada
        /// </summary>
        /// <param name="rutaDirectorio">contiene la rutaDirectorio donde generar los ficheros. Hay que concatenarles el nombre del fichero</param>
        /// <param name="ofertasSeleccionadas">listado de las ofertas que el usuario ha seleccionado con un Check en la interfaz.</param>
        public ResultBE GenerarFichasCorreosExpress(string rutaDirectorio, Collection<OfertaChronoFichaBE> ofertasSeleccionadas)
        {
            ResultBE objRespuesta = new ResultBE();

            StringBuilder sbError = new StringBuilder();
            bool procesoOK = true;

            foreach (OfertaChronoFichaBE objOferta in ofertasSeleccionadas)
            {
                bool ficheroOK = true;

                //Obtenemos el dichero plantilla y el fichero destino
                string FicheroOrigen = string.Format(Utils.GetValorFromAppConfig(AppSettingsEnum.RutaFichaCEX), AppDomain.CurrentDomain.BaseDirectory);
                string FicheroDestino = Path.Combine(rutaDirectorio, string.Format(CultureInfo.InvariantCulture, "{0}{1}{2}{1}{3}{4}{5}", objOferta.CodOfertaSAP, "_", objOferta.Cliente.CodClienteSAP, DateTime.Now.Day.ToString(), string.Format(CultureInfo.InvariantCulture, "{0:00}", DateTime.Now.Month), DateTime.Now.Year.ToString()));

                Dictionary<string, object> objContenidoARellenarInforme = new Dictionary<string, object>();
                Dictionary<string, object> objContenidoARellenarAsistente = new Dictionary<string, object>();
                Dictionary<string, object> objContenidoARellenarBalearesCanarias = new Dictionary<string, object>();

                //FicheroDestino que se modificará para generar el informe
                ManagerExcel objExcel = new ManagerExcel(FicheroOrigen, false);

                try
                {
                    objExcel.AbrirFichero();
                    if (objExcel.Abierto)
                    {
                        //Remplazar contenido
                        #region Pestaña informe

                        objExcel.SeleccionarHoja(SimuladorResources.PestanaInformeFichaCex);
                        objExcel.DesprotegerHoja("3030");

                        //Nombre cliente 
                        if (!string.IsNullOrWhiteSpace(objOferta.NombreClienteListado))
                        {
                            objExcel.EscribirEnCeldaEditable("C5", "I5", objOferta.NombreClienteListado);
                        }

                        if (objOferta.Cliente != null)
                        {
                            //CIF cliente 
                            if (!string.IsNullOrWhiteSpace(objOferta.Cliente.CIF))
                            {
                                objExcel.EscribirEnCeldaProteger("Q5", "R5", objOferta.Cliente.CIF);
                            }

                            //Código Postal cliente
                            if (!string.IsNullOrWhiteSpace(objOferta.Cliente.CP))
                            {
                                objExcel.EscribirEnCeldaProteger("C7", objOferta.Cliente.CP);
                            }

                            //Provincia cliente 
                            if (!string.IsNullOrWhiteSpace(objOferta.Cliente.Provincia))
                            {
                                objExcel.EscribirEnCeldaProteger("E7", "F7", objOferta.Cliente.Provincia);
                            }

                            //Población cliente  
                            if (!string.IsNullOrWhiteSpace(objOferta.Cliente.Ciudad))
                            {
                                objExcel.EscribirEnCeldaProteger("H7", "I7", objOferta.Cliente.Ciudad);
                            }
                        }

                        //Fecha inicio oferta
                        objExcel.EscribirEnCeldaProteger("O2", "R2", objOferta.FechaInicioExcel);

                        //Se protege la celda M5, fecha desde
                        objExcel.ProtegerCelda("M5");

                        //Fecha fin oferta  
                        if (!string.IsNullOrWhiteSpace(objOferta.FechaFinExcel))
                        {
                            objExcel.EscribirEnCeldaProteger("M6", "M7", objOferta.FechaFinExcel);
                        }

                        #region Valores Añadidos

                        #region Reembolso

                        PosicionSAPBE posicion = objOferta.ListaPosiciones.FirstOrDefault(x => !x.PosicionReferencia.Equals(0) && x.Producto.Equals(SimuladorResources.SAC141));
                        if (posicion != null)
                        {
                            objExcel.EscribirEnCeldaProteger("R8", SimuladorResources.SI);

                            //Se obtiene el nodo descuento
                            DestinoSAPBE destinoReembolso = objOferta.ListaDestinos.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x.CodValorAnadidoSAP) && x.CodValorAnadidoSAP.Equals(SimuladorResources.SAC141));

                            if (destinoReembolso != null && destinoReembolso.DescuentoFinal > 0)
                            {
                                decimal auxPorcentaje = 0, auxImporte = 0;

                                //Se lee el valor previo de la celda %
                                string celdaPorcentaje = objExcel.LeerDeCelda("R9");

                                //Se lee el valor previo de la celda €
                                string celdaImporte = objExcel.LeerDeCelda("R10");

                                if (decimal.TryParse(celdaPorcentaje, out auxPorcentaje) && decimal.TryParse(celdaImporte, out auxImporte))
                                {
                                    //Se calculan los nuevos valores
                                    auxPorcentaje = auxPorcentaje - (auxPorcentaje * (destinoReembolso.DescuentoFinal / 100));
                                    auxImporte = auxImporte - (auxImporte * (destinoReembolso.DescuentoFinal / 100));

                                    //Se escriben los datos
                                    objExcel.EscribirEnCeldaProteger("R9", auxPorcentaje);
                                    objExcel.EscribirEnCeldaProteger("R10", auxImporte);
                                }
                            }
                        }

                        #endregion

                        #region Tipo Seguro

                        posicion = objOferta.ListaPosiciones.FirstOrDefault(x => !x.PosicionReferencia.Equals(0) && x.Producto.Equals(SimuladorResources.SAC143));

                        if (posicion != null)
                        {
                            objExcel.EscribirEnCeldaProteger("R11", SimuladorResources.TodoRiesgo);
                        }

                        #endregion

                        #region Entrega franja horaria

                        posicion = objOferta.ListaPosiciones.FirstOrDefault(x => !x.PosicionReferencia.Equals(0) && x.Producto.Equals(SimuladorResources.SAC146));

                        if (posicion != null)
                        {
                            //Se obtiene el nodo descuento
                            DestinoSAPBE destinoReembolso = objOferta.ListaDestinos.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x.CodValorAnadidoSAP) && x.CodValorAnadidoSAP.Equals(SimuladorResources.SAC146));

                            if (destinoReembolso != null && destinoReembolso.DescuentoFinal > 0)
                            {
                                decimal auxImporte = 0;

                                //Se lee el valor previo de la celda importe
                                string celdaImporte = objExcel.LeerDeCelda("R12");

                                if (decimal.TryParse(celdaImporte, out auxImporte))
                                {
                                    //Se calculan los nuevos valores
                                    auxImporte = auxImporte - (auxImporte * (destinoReembolso.DescuentoFinal / 100));

                                    //Se escribe el dato
                                    objExcel.EscribirEnCeldaProteger("R12", auxImporte);
                                }
                            }
                        }

                        #endregion

                        #region Entrega en sábado

                        posicion = objOferta.ListaPosiciones.FirstOrDefault(x => !x.PosicionReferencia.Equals(0) && x.Producto.Equals(SimuladorResources.SAC140));

                        if (posicion != null)
                        {
                            //Se obtiene el nodo descuento
                            DestinoSAPBE destinoReembolso = objOferta.ListaDestinos.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x.CodValorAnadidoSAP) && x.CodValorAnadidoSAP.Equals(SimuladorResources.SAC140));

                            if (destinoReembolso != null && destinoReembolso.DescuentoFinal > 0)
                            {
                                decimal auxImporte = 0;

                                //Se lee el valor previo de la celda importe
                                string celdaImporte = objExcel.LeerDeCelda("R13");

                                if (decimal.TryParse(celdaImporte, out auxImporte))
                                {
                                    //Se calculan los nuevos valores
                                    auxImporte = auxImporte - (auxImporte * (destinoReembolso.DescuentoFinal / 100));

                                    //Se escribe el dato
                                    objExcel.EscribirEnCeldaProteger("R13", auxImporte);
                                }
                            }
                        }

                        #endregion

                        #region Retorno albarán cliente

                        posicion = objOferta.ListaPosiciones.FirstOrDefault(x => !x.PosicionReferencia.Equals(0) && x.Producto.Equals(SimuladorResources.SAC145));

                        if (posicion != null)
                        {
                            //Se obtiene el nodo descuento
                            DestinoSAPBE destinoReembolso = objOferta.ListaDestinos.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x.CodValorAnadidoSAP) && x.CodValorAnadidoSAP.Equals(SimuladorResources.SAC145));

                            if (destinoReembolso != null && destinoReembolso.DescuentoFinal > 0)
                            {
                                decimal auxImporte = 0;

                                //Se lee el valor previo de la celda importe
                                string celdaImporte = objExcel.LeerDeCelda("R14");

                                if (decimal.TryParse(celdaImporte, out auxImporte))
                                {
                                    //Se calculan los nuevos valores
                                    auxImporte = auxImporte - (auxImporte * (destinoReembolso.DescuentoFinal / 100));

                                    //Se escribe el dato
                                    objExcel.EscribirEnCeldaProteger("R14", auxImporte);
                                }
                            }
                        }

                        #endregion

                        #region SMS

                        //El SMS, código S0287, aunque es un valor añadido, SAP lo gestiona como si fuera un producto
                        posicion = objOferta.ListaPosiciones.FirstOrDefault(x => x.PosicionReferencia.Equals(0) && x.Producto.Equals(SimuladorResources.S0287));

                        if (posicion != null)
                        {
                            //Se obtiene el nodo descuento
                            DestinoSAPBE destinoReembolso = objOferta.ListaDestinos.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x.CodProductoSAP) && x.CodProductoSAP.Equals(SimuladorResources.S0287));

                            if (destinoReembolso != null && destinoReembolso.DescuentoFinal > 0)
                            {
                                decimal auxImporte = 0;

                                //Se lee el valor previo de la celda importe
                                string celdaImporte = objExcel.LeerDeCelda("R15");

                                if (decimal.TryParse(celdaImporte, out auxImporte))
                                {
                                    //Se calculan los nuevos valores
                                    auxImporte = auxImporte - (auxImporte * (destinoReembolso.DescuentoFinal / 100));

                                    //Se escribe el dato
                                    objExcel.EscribirEnCeldaProteger("R15", auxImporte);
                                }
                            }
                        }

                        #endregion

                        #endregion

                        //Zona nacional no lo tenemos                        
                        //objContenidoARellenarInforme.Add("C8", "1"); //Zona nacional (1, 2, 3, 4, 5, 6, 7) 

                        if (objOferta.DatosGestor != null)
                        {
                            objExcel.EscribirEnCeldaProteger("C9", "F9", objOferta.DatosGestor.FullName); //Comercial 
                            objExcel.EscribirEnCeldaProteger("C10", objOferta.DatosGestor.TelefonoUsuario); //Teléfono RICO 

                            string[] mail = objOferta.DatosGestor.EmailUsuario.Split('@');
                            if (mail.Length > 0)
                            {
                                objExcel.EscribirEnCeldaProteger("C11", "E11", mail[0]); //Dirección de email sin dominio 
                            }
                        }

                        //El Código Comercial CEX no lo tenemos
                        //objContenidoARellenarInforme.Add("F10", "Cod CEX"); //Código Comercial CEX                         

                        //Facturación Mensual Neta
                        objExcel.EscribirEnCeldaProteger("D13", "E14", objOferta.FacturacionNeta);

                        //Condición y formas de pago
                        if (objOferta.Cliente != null)
                        {
                            //Solo se establece la forma de pago Transferencia y Cheque
                            //Para el resto se dejará la opción por defect: Domiciliación

                            if (!string.IsNullOrWhiteSpace(objOferta.Cliente.CodFormaPago) &&
                                objOferta.Cliente.CodFormaPago.Equals(SimuladorResources.TransferenciaCodigoSAP))
                            {
                                //Transferencia
                                objExcel.EscribirEnCeldaProteger("L35", "M35", SimuladorResources.Transferencia);
                            }
                            else if (!string.IsNullOrWhiteSpace(objOferta.Cliente.CodFormaPago) &&
                                objOferta.Cliente.CodFormaPago.Equals(SimuladorResources.ChequeTalonCodigoSAP))
                            {
                                //Cheque
                                objExcel.EscribirEnCeldaProteger("L35", "M35", SimuladorResources.ModoCobroChequeTalon);
                            }

                            if (!string.IsNullOrWhiteSpace(objOferta.Cliente.CodCondicionPago) &&
                                objOferta.Cliente.CodCondicionPago.Equals(SimuladorResources.Vencimiento10Dias))
                            {
                                //Vencimiento a 10 días
                                objExcel.EscribirEnCeldaProteger("L36", "M36", SimuladorResources.Valor10Dias);
                            }
                            else if (!string.IsNullOrWhiteSpace(objOferta.Cliente.CodCondicionPago) &&
                                objOferta.Cliente.CodCondicionPago.Equals(SimuladorResources.Vencimiento30Dias))
                            {
                                //Vencimiento a 30 días
                                objExcel.EscribirEnCeldaProteger("L36", "M36", SimuladorResources.Valor30Dias);
                            }
                        }

                        objExcel.ProtegerHoja("3030");

                        #endregion

                        #region Pestaña Asistente

                        #region Paq10

                        List<DestinoSAPBE> destinos = objOferta.ListaDestinos.Where(x => x.CodProductoSAP.Equals(SimuladorResources.S0280)).ToList<DestinoSAPBE>();

                        if (destinos != null && destinos.Count > 0)
                        {
                            //Zonas
                            //Provincial
                            DestinoSAPBE destino = destinos.FirstOrDefault(x => x.Destino.Equals("Z01"));

                            if (destino != null && destino.DescuentoFinal > 0)
                            {
                                objContenidoARellenarAsistente.Add("C13", destino.DescuentoFinal / 100);
                            }

                            //Regional
                            destino = destinos.FirstOrDefault(x => x.Destino.Equals("Z02"));

                            if (destino != null && destino.DescuentoFinal > 0)
                            {
                                objContenidoARellenarAsistente.Add("D13", destino.DescuentoFinal / 100);
                            }

                            //Nacional
                            destino = destinos.FirstOrDefault(x => x.Destino.Equals("Z03"));

                            if (destino != null && destino.DescuentoFinal > 0)
                            {
                                objContenidoARellenarAsistente.Add("E13", destino.DescuentoFinal / 100);
                            }

                            //Nacional+
                            destino = destinos.FirstOrDefault(x => x.Destino.Equals("Z11"));

                            if (destino != null && destino.DescuentoFinal > 0)
                            {
                                objContenidoARellenarAsistente.Add("F13", destino.DescuentoFinal / 100);
                            }

                            //Mallorca
                            destino = destinos.FirstOrDefault(x => x.Destino.Equals("Z04"));

                            if (destino != null && destino.DescuentoFinal > 0)
                            {
                                objContenidoARellenarAsistente.Add("H13", destino.DescuentoFinal / 100);
                            }

                            //Islas Menores
                            destino = destinos.FirstOrDefault(x => x.Destino.Equals("Z14"));

                            if (destino != null && destino.DescuentoFinal > 0)
                            {
                                objContenidoARellenarAsistente.Add("I13", destino.DescuentoFinal / 100);
                            }

                            //Tnf y Lpa
                            destino = destinos.FirstOrDefault(x => x.Destino.Equals("Z05"));

                            if (destino != null && destino.DescuentoFinal > 0)
                            {
                                objContenidoARellenarAsistente.Add("K13", destino.DescuentoFinal / 100);
                            }
                        }

                        #endregion

                        #region Paq14

                        destinos = objOferta.ListaDestinos.Where(x => x.CodProductoSAP.Equals(SimuladorResources.S0281)).ToList<DestinoSAPBE>();

                        if (destinos != null && destinos.Count > 0)
                        {
                            //Zonas
                            //Provincial
                            DestinoSAPBE destino = destinos.FirstOrDefault(x => x.Destino.Equals("Z01"));

                            if (destino != null && destino.DescuentoFinal > 0)
                            {
                                objContenidoARellenarAsistente.Add("C14", destino.DescuentoFinal / 100);
                            }

                            //Regional
                            destino = destinos.FirstOrDefault(x => x.Destino.Equals("Z02"));

                            if (destino != null && destino.DescuentoFinal > 0)
                            {
                                objContenidoARellenarAsistente.Add("D14", destino.DescuentoFinal / 100);
                            }

                            //Nacional
                            destino = destinos.FirstOrDefault(x => x.Destino.Equals("Z03"));

                            if (destino != null && destino.DescuentoFinal > 0)
                            {
                                objContenidoARellenarAsistente.Add("E14", destino.DescuentoFinal / 100);
                            }

                            //Nacional+
                            destino = destinos.FirstOrDefault(x => x.Destino.Equals("Z11"));

                            if (destino != null && destino.DescuentoFinal > 0)
                            {
                                objContenidoARellenarAsistente.Add("F14", destino.DescuentoFinal / 100);
                            }

                            //Mallorca
                            destino = destinos.FirstOrDefault(x => x.Destino.Equals("Z04"));

                            if (destino != null && destino.DescuentoFinal > 0)
                            {
                                objContenidoARellenarAsistente.Add("H14", destino.DescuentoFinal / 100);
                            }

                            //Islas Menores
                            destino = destinos.FirstOrDefault(x => x.Destino.Equals("Z14"));

                            if (destino != null && destino.DescuentoFinal > 0)
                            {
                                objContenidoARellenarAsistente.Add("I14", destino.DescuentoFinal / 100);
                            }

                            //Tnf y Lpa
                            destino = destinos.FirstOrDefault(x => x.Destino.Equals("Z05"));

                            if (destino != null && destino.DescuentoFinal > 0)
                            {
                                objContenidoARellenarAsistente.Add("K14", destino.DescuentoFinal / 100);
                            }
                        }

                        #endregion

                        #region Paq24

                        destinos = objOferta.ListaDestinos.Where(x => x.CodProductoSAP.Equals(SimuladorResources.S0282)).ToList<DestinoSAPBE>();

                        if (destinos != null && destinos.Count > 0)
                        {
                            //Zonas
                            //Provincial
                            DestinoSAPBE destino = destinos.FirstOrDefault(x => x.Destino.Equals("Z01"));

                            if (destino != null && destino.DescuentoFinal > 0)
                            {
                                objContenidoARellenarAsistente.Add("C15", destino.DescuentoFinal / 100);
                            }

                            //Regional
                            destino = destinos.FirstOrDefault(x => x.Destino.Equals("Z02"));

                            if (destino != null && destino.DescuentoFinal > 0)
                            {
                                objContenidoARellenarAsistente.Add("D15", destino.DescuentoFinal / 100);
                            }

                            //Nacional
                            destino = destinos.FirstOrDefault(x => x.Destino.Equals("Z03"));

                            if (destino != null && destino.DescuentoFinal > 0)
                            {
                                objContenidoARellenarAsistente.Add("E15", destino.DescuentoFinal / 100);
                            }

                            //Nacional+
                            destino = destinos.FirstOrDefault(x => x.Destino.Equals("Z11"));

                            if (destino != null && destino.DescuentoFinal > 0)
                            {
                                objContenidoARellenarAsistente.Add("F15", destino.DescuentoFinal / 100);
                            }

                            //Baleares Interislas
                            destino = destinos.FirstOrDefault(x => x.Destino.Equals("Z15"));

                            if (destino != null && destino.DescuentoFinal > 0)
                            {
                                objContenidoARellenarAsistente.Add("J15", destino.DescuentoFinal / 100);
                            }

                            //Canarias Interislas
                            destino = destinos.FirstOrDefault(x => x.Destino.Equals("Z06"));

                            if (destino != null && destino.DescuentoFinal > 0)
                            {
                                objContenidoARellenarAsistente.Add("M15", destino.DescuentoFinal / 100);
                            }

                            //Ceuta y Melilla
                            destino = destinos.FirstOrDefault(x => x.Destino.Equals("Z16"));

                            if (destino != null && destino.DescuentoFinal > 0)
                            {
                                objContenidoARellenarAsistente.Add("N15", destino.DescuentoFinal / 100);
                            }

                            //Andorra
                            destino = destinos.FirstOrDefault(x => x.Destino.Equals("Z13"));

                            if (destino != null && destino.DescuentoFinal > 0)
                            {
                                objContenidoARellenarAsistente.Add("O15", destino.DescuentoFinal / 100);
                            }

                            //Gibraltar
                            destino = destinos.FirstOrDefault(x => x.Destino.Equals("Z12"));

                            if (destino != null && destino.DescuentoFinal > 0)
                            {
                                objContenidoARellenarAsistente.Add("P15", destino.DescuentoFinal / 100);
                            }

                            //Portugal Peninsular
                            destino = destinos.FirstOrDefault(x => x.Destino.Equals("Z07"));

                            if (destino != null && destino.DescuentoFinal > 0)
                            {
                                objContenidoARellenarAsistente.Add("R15", destino.DescuentoFinal / 100);
                            }

                            //Portugal Islas
                            destino = destinos.FirstOrDefault(x => x.Destino.Equals("Z10"));

                            if (destino != null && destino.DescuentoFinal > 0)
                            {
                                objContenidoARellenarAsistente.Add("S15", destino.DescuentoFinal / 100);
                            }
                        }

                        #endregion

                        #endregion

                        #region Pestaña Baleares y Canarias

                        #region Baleares

                        destinos = objOferta.ListaDestinos.Where(x => x.CodProductoSAP.Equals("S0283")).ToList<DestinoSAPBE>();

                        if (destinos != null && destinos.Count > 0)
                        {
                            //Zonas
                            //Mallorca
                            DestinoSAPBE destino = destinos.FirstOrDefault(x => x.Destino.Equals("Z04"));

                            if (destino != null && destino.DescuentoFinal > 0)
                            {
                                objContenidoARellenarBalearesCanarias.Add("C18", destino.DescuentoFinal / 100);
                            }

                            //Baleares Islas Menores
                            destino = destinos.FirstOrDefault(x => x.Destino.Equals("Z14"));

                            if (destino != null && destino.DescuentoFinal > 0)
                            {
                                objContenidoARellenarBalearesCanarias.Add("D18", destino.DescuentoFinal / 100);
                            }
                        }

                        #endregion

                        #region Canarias Express

                        destinos = objOferta.ListaDestinos.Where(x => x.CodProductoSAP.Equals("S0284")).ToList<DestinoSAPBE>();

                        if (destinos != null && destinos.Count > 0)
                        {
                            //Zonas
                            //Tenerife y Las Palmas
                            DestinoSAPBE destino = destinos.FirstOrDefault(x => x.Destino.Equals("Z05"));

                            if (destino != null && destino.DescuentoFinal > 0)
                            {
                                objContenidoARellenarBalearesCanarias.Add("G18", destino.DescuentoFinal / 100);
                            }

                            //Canarias Islas Menores
                            destino = destinos.FirstOrDefault(x => x.Destino.Equals("Z17"));

                            if (destino != null && destino.DescuentoFinal > 0)
                            {
                                objContenidoARellenarBalearesCanarias.Add("H18", destino.DescuentoFinal / 100);
                            }
                        }

                        #endregion

                        #region Canarias Aéreo

                        destinos = objOferta.ListaDestinos.Where(x => x.CodProductoSAP.Equals("S0285")).ToList<DestinoSAPBE>();

                        if (destinos != null && destinos.Count > 0)
                        {
                            //Zonas
                            //Tenerife y Las Palmas
                            DestinoSAPBE destino = destinos.FirstOrDefault(x => x.Destino.Equals("Z05"));

                            if (destino != null && destino.DescuentoFinal > 0)
                            {
                                objContenidoARellenarBalearesCanarias.Add("G19", destino.DescuentoFinal / 100);
                            }

                            //Canarias Islas Menores
                            destino = destinos.FirstOrDefault(x => x.Destino.Equals("Z17"));

                            if (destino != null && destino.DescuentoFinal > 0)
                            {
                                objContenidoARellenarBalearesCanarias.Add("H19", destino.DescuentoFinal / 100);
                            }
                        }

                        #endregion

                        #region Canarias Marítimo

                        destinos = objOferta.ListaDestinos.Where(x => x.CodProductoSAP.Equals("S0286")).ToList<DestinoSAPBE>();

                        if (destinos != null && destinos.Count > 0)
                        {
                            //Zonas
                            //Tenerife y Las Palmas
                            DestinoSAPBE destino = destinos.FirstOrDefault(x => x.Destino.Equals("Z05"));

                            if (destino != null && destino.DescuentoFinal > 0)
                            {
                                objContenidoARellenarBalearesCanarias.Add("G20", destino.DescuentoFinal / 100);
                            }

                            //Canarias Islas Menores
                            destino = destinos.FirstOrDefault(x => x.Destino.Equals("Z17"));

                            if (destino != null && destino.DescuentoFinal > 0)
                            {
                                objContenidoARellenarBalearesCanarias.Add("H20", destino.DescuentoFinal / 100);
                            }
                        }

                        #endregion

                        #endregion

                        //Se escribe los datos en las pestañas correspondientes
                        this.EscribirDatosPestañaFichaCEX(objExcel, SimuladorResources.PestanaAsistenteFichaCex, objContenidoARellenarAsistente);
                        this.EscribirDatosPestañaFichaCEX(objExcel, SimuladorResources.PestanaBalearesCanariasFichaCex, objContenidoARellenarBalearesCanarias);

                        objExcel.GuardarComoLibro(FicheroDestino);
                        objExcel.CerrarExcel();
                    }
                }
                catch (Exception ex)
                {
                    if (objExcel.Abierto)
                    {
                        objExcel.CerrarExcel();
                    }

                    RegistrarAccionesSimulador.GuardarExcepcion(ex, false);
                    ficheroOK = false;
                }
                finally
                {
                    if (objExcel.Abierto)
                    {
                        objExcel.CerrarExcel();
                    }
                    if (!ficheroOK)
                    {
                        sbError.AppendLine(string.Format(CultureInfo.InvariantCulture, "{0}", objOferta.CodOfertaSAP));
                    }
                    procesoOK = procesoOK && ficheroOK;
                }
            }
            if (!procesoOK)
            {
                objRespuesta.Resultado = procesoOK;
                objRespuesta.TextoError = string.Format(CultureInfo.InvariantCulture, "{0}{1}{2}", SimuladorResources.ErrorGenerarFichasCEX, Environment.NewLine, sbError.ToString());
            }
            return objRespuesta;
        }
        #endregion

        #endregion

        #endregion

        #region Métodos Privados

        #region Escribir Pestañas Ficha CEX

        /// <summary>
        /// Método que escribe en la pestaña correspondiente los datos pasados por parámetro
        /// </summary>
        /// <param name="objExcel">Objeto excel</param>
        /// <param name="nombrePestaña">Nombre de la pestaña donde se van a escribir los datos</param>
        /// <param name="datos">Diccionario de datos a escribir</param>
        private void EscribirDatosPestañaFichaCEX(ManagerExcel objExcel, string nombrePestaña, Dictionary<string, object> datos)
        {
            objExcel.SeleccionarHoja(nombrePestaña);
            objExcel.DesprotegerHoja("3030");
            objExcel.EscribirCeldasProtegidasMultiple(datos);
            objExcel.ProtegerHoja("3030");
        }

        #endregion

        #region CrearTablaVA

        /// <summary>
        /// Método que devuelve la lista de tarifas de los valores añadidos
        /// </summary>
        /// <param name="idProducto">Identificador del que se desean obtener sus productos</param>
        /// <param name="idProductoOferta">Identificador del producto oferta</param>
        /// <returns>Lista de valores añadidos</returns>
        private Collection<ReporteVABE> CrearTablaVA(Guid idProducto, Guid idProductoOferta, bool tarifasConDescuento)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ValorAnadidoProductoPersistence persistence = new ValorAnadidoProductoPersistence(uow);
                return persistence.ObtenerReporteTarifasVA(idProducto, idProductoOferta, tarifasConDescuento);
            }
        }

        #endregion

        #region Informes

        /// <summary>
        /// Obtiene la descripción del tramo con formato Desde _ Hasta _ Kg/g 
        /// </summary>
        /// <param name="objProductoBE"></param>
        /// <param name="objTramo"></param>
        /// <param name="numTramos"></param>
        /// <param name="indiceTramo"></param>
        /// <param name="pesoTramo"></param>
        /// <returns></returns>
        private String GetDescripionTramo(ProductoBE objProductoBE, TramoBE objTramo, List<TramoInformeBE> listaTramosInfTarifas, int numTramos, int indiceTramo, ref Decimal pesoTramo)
        {
            String descripcionTramo = String.Empty;

            //Paqueteria kg, si no en g. Miramos si su modelo de descuento es paquetería, tramos, o aparece como paquete en el doc. de tarifas
            //Si es Publicorreo óptimo, no lo consideramos como paquetería     
            bool paqueteria = !ModeloDescuentoEnum.GetPaqueteriasQueSonPublicorreos().Contains(objProductoBE.CodProducto) && objProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Paqueteria)) ||
                              objProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Tramos)) ||
                              listaTramosInfTarifas.Count > 0;


            //Si es paquetería se muestra el texto Más De __ Kg hasta __ Kg
            if (paqueteria && !objProductoBE.CodProducto.Equals("S0360"))
            {
                Func<decimal, string> formatearKg = (numDec => (numDec % 1) != 0 ? numDec.ToString("N3") : numDec.ToString("N3").TrimEnd('0').TrimEnd(','));

                //if (indiceTramo == 0 && objTramo.RangoDeTramo.Item1 == 0)
                if (objTramo.RangoDeTramo.Item1 == 0)
                {
                    pesoTramo = ((decimal)objTramo.RangoDeTramo.Item2) / 1000;
                    descripcionTramo = "Hasta " + formatearKg(pesoTramo) + " Kg.";
                }
                else if (indiceTramo < numTramos)// && objTramo.RangoDeTramo.Item2 != 30000)
                {
                    pesoTramo = Math.Round(((decimal)objTramo.RangoDeTramo.Item1) / 10, MidpointRounding.AwayFromZero) * 10 / 1000;
                    descripcionTramo = "Más de " + formatearKg(pesoTramo) + " hasta ";
                    pesoTramo = ((decimal)objTramo.RangoDeTramo.Item2) / 1000;
                    descripcionTramo += formatearKg(pesoTramo) + " Kg.";
                }              
                //else
                //{

                //    descripcionTramo = "Kg. Adicional hasta 30 Kg.";
                //}

            }
            else if (objProductoBE.CodProducto.Equals("S0360"))
            {
                //if (indiceTramo == 0 && objTramo.RangoDeTramo.Item1 == 0)
                if (objTramo.RangoDeTramo.Item1 == 0)
                {
                    pesoTramo = ((decimal)objTramo.RangoDeTramo.Item2);
                    descripcionTramo = "Hasta " + pesoTramo + " g.";
                }
                else if (indiceTramo < numTramos)
                {
                    if (objTramo.RangoDeTramo.Item2 != null)
                    {
                        var tramoRedondeado = Math.Round(((decimal)objTramo.RangoDeTramo.Item1) / 10, MidpointRounding.AwayFromZero) * 10;
                        descripcionTramo = "Más de " + tramoRedondeado + " hasta ";
                        pesoTramo = ((decimal)objTramo.RangoDeTramo.Item2);
                        descripcionTramo += pesoTramo + " g.";
                    }
                    else
                    {
                        var tramoRedondeado = Math.Round(((decimal)objTramo.RangoDeTramo.Item1) / 10, MidpointRounding.AwayFromZero) * 10;
                        descripcionTramo = tramoRedondeado + " g.";
                    }
                }
            } 
            //Si no es Paquetería, se muestra en Gramos.
            else
            {
                //if (indiceTramo == 0 && objTramo.RangoDeTramo.Item1 == 0)
                if (objTramo.RangoDeTramo.Item1 == 0)
                {
                    pesoTramo = ((decimal)objTramo.RangoDeTramo.Item2);
                    descripcionTramo = "Hasta " + pesoTramo + " g.";
                }
                else if (indiceTramo < numTramos)
                {
                    if (objTramo.RangoDeTramo.Item2 != null)
                    {
                        pesoTramo = ((decimal)objTramo.RangoDeTramo.Item1);
                        descripcionTramo = "Más de " + pesoTramo + " hasta ";
                        pesoTramo = ((decimal)objTramo.RangoDeTramo.Item2);
                        descripcionTramo += pesoTramo + " g.";
                    }
                    else
                    {
                        pesoTramo = ((decimal)objTramo.RangoDeTramo.Item1);
                        descripcionTramo = pesoTramo + " g.";
                    }
                }
            }

            if (objTramo.Descripcion.Contains("N"))
                descripcionTramo += " normalizadas";

            descripcionTramo = descripcionTramo + '\t';

            return descripcionTramo;
        }

        /// <summary>
        /// Obtiene la descripción del tramo con formato Desde _ Hasta _ Kg/g 
        /// </summary>
        /// <param name="objProductoBE"></param>
        /// <param name="objTramo"></param>
        /// <param name="numTramos"></param>
        /// <param name="indiceTramo"></param>
        /// <param name="pesoTramo"></param>
        /// <returns></returns>
        private String GetDescripionGrupoTramo(ProductoBE objProductoBE, GrupoTramoBE objTramo, List<TramoInformeBE> listaTramosInfTarifas, int numTramos, int indiceTramo, bool esPublicorreo)
        {
            String descripcionTramo = String.Empty, unidad;
            String pesoTramoIni, pesoTramoFin;
            int decTramoIni = 0, decTramoFin = 0;

            //Paqueteria kg, si no en g. Miramos si su modelo de descuento es paquetería, tramos, o aparece como paquete en el doc. de tarifas                
            //Si es Publicorreo óptimo, no lo consideramos como paquetería     
            bool esPaqueteria = !ModeloDescuentoEnum.GetPaqueteriasQueSonPublicorreos().Contains(objProductoBE.CodProducto) && objProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Paqueteria)) ||
                              objProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Tramos)) ||
                              listaTramosInfTarifas.Count > 0;

            if (objTramo.ListaTramos != null && objTramo.ListaTramos.Any())
            {
                decTramoIni = objTramo.ListaTramos.First().RangoDeTramo.Item1 ?? 0;

                if(esPublicorreo)
                    decTramoFin = objTramo.ListaTramos.Last().RangoDeTramo.Item1 ?? 0;
                else
                    decTramoFin = objTramo.ListaTramos.Last().RangoDeTramo.Item2 ?? 0;
            }
            
            //Si es paquetería se muestra el texto Más De __ Kg hasta __ Kg
            if (esPaqueteria && !objProductoBE.CodProducto.Equals("S0360"))
            {
                Func<double, string> formatearKg = (numDec => (numDec % 1) != 0 ? numDec.ToString("N3") : numDec.ToString("N3").TrimEnd('0').TrimEnd(','));

                var tramoRedondeado = Math.Round(((decimal)decTramoIni) / 10, MidpointRounding.AwayFromZero) * 10;
                pesoTramoIni = formatearKg((double)tramoRedondeado / 1000.0);
                pesoTramoFin = formatearKg(decTramoFin / 1000.0);
                unidad = " Kg.";
            }
            //Si no es Paquetería, se muestra en Gramos.
            else if (objProductoBE.CodProducto.Equals("S0360"))
            {
                var tramoRedondeado = Math.Round(((decimal)decTramoIni) / 10, MidpointRounding.AwayFromZero) * 10;
                unidad = " g.";
                pesoTramoIni = tramoRedondeado.ToString();
                pesoTramoFin = decTramoFin.ToString();

            } 
            else 
            {
                unidad = " g.";
                pesoTramoIni = decTramoIni.ToString();
                pesoTramoFin = decTramoFin.ToString();
            }

            descripcionTramo = ((objTramo.Nombre.Contains("Hasta")) ? "Hasta " + pesoTramoFin.Trim() : "Más de " + pesoTramoIni.Trim() + " hasta " + pesoTramoFin) + unidad;

            return descripcionTramo;
        }

        #endregion

        private string GenerarAbreviaturaModeloNegociacion(string CodModalidadNegociacion) 
        {
            switch (CodModalidadNegociacion.Trim())
            {
                case "5DD":
                    return "DD";
                case "4DT":
                    return "DT";
                case "2PT":
                    return "PC";
                case "3DG":
                    return "DGT";
                case "1PG":
                    return "PCGT";
                default:
                    return "";

                #region "Por descripcion"

                //case "Descuento por destino":
                //    return "DD";
                //case "Descuento por tramo":
                //    return "DT";
                //case "Precio Cierto por tramo":
                //    return "PCT";
                //case "Descuento por Grupo de Tramo":
                //    return "DGT";
                //case "Precio Cierto por Grupo de Tramo":
                //    return "PCGT";
                //default:
                //    return "";

                #endregion
            }


        }

        #endregion
    }
}
