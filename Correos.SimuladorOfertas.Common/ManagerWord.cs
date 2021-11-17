using System;
using System.Collections.Generic;
using System.Globalization;
using Correos.SimuladorOfertas.Common.Enums;
using Microsoft.Office.Interop.Word;
using System.Drawing;
using System.Configuration;

namespace Correos.SimuladorOfertas.Common
{
    public class ManagerWord
    {
        #region Miembros de la pagina

        // Aplicacion excel
        public AplicacionWord word;
        // Documentos
        protected Microsoft.Office.Interop.Word.Documents documentosWord;
        protected Microsoft.Office.Interop.Word.Document documento;        

        protected int numTablaSeleccionada = 0;

        // Para creacion de tablas
        //protected Microsoft.Office.Interop.Word.Table laTablaCreada;

        object m_objOpt;

        // Nombre del fichero
        protected string nomFich;

        // Indica si hay un fichero excel abierto
        protected bool abierto;

        // Fin del documento        
        public static object END_OF_DOC = "\\endofdoc";

        #endregion

        #region Propiedades

        /// <summary>
        /// Variable para ver la aplicación Word.
        /// </summary>
        public bool VerWord
        {
            get { return this.word.Word.Visible; }
            set { this.word.Word.Visible = value; }
        }

        /// <summary>
        /// Devuelve si el fichero doc está abierto.
        /// </summary>
        public bool Abierto
        {
            get { return this.abierto; }
        }

        /// <summary>
        /// Nombre del fichero
        /// </summary>
        public string Fichero
        {
            get { return this.nomFich; }
            set { this.nomFich = value; }
        }

        #endregion

        #region Constructores

        public ManagerWord(string nF, bool visible)
        {
            this.nomFich = nF;

            this.word = new AplicacionWord();
            this.word.Word.Visible = visible;


#if DEBUG
            if (ConfigurationManager.AppSettings["OfficeAppVisible"] == "S")
            {
                this.word.Word.Visible = true;
            }
#endif

            this.m_objOpt = System.Reflection.Missing.Value;
        }

        public ManagerWord()
        {
            this.nomFich = null;

            this.word = new AplicacionWord();
            this.word.Word.Visible = false;

            this.m_objOpt = System.Reflection.Missing.Value;
        }



        #endregion

        #region Métodos Públicos

        #region Apertura/Creación del fichero

        public void DesactivarModoWord()
        {
            this.word.Word.Application.Visible = true;
            this.word.Word.Application.WindowActivate -= new ApplicationEvents4_WindowActivateEventHandler(evento);
        }

        public void ActivarModoWord()
        {
            this.word.Word.Application.WindowActivate += new ApplicationEvents4_WindowActivateEventHandler(evento);
        }

        private void evento(Document Doc, Window Wn)
        {
            this.word.Word.Application.Visible = false;
#if DEBUG
            if (ConfigurationManager.AppSettings["OfficeAppVisible"] == "S")
            {
                this.word.Word.Visible = true;
            }
#endif

        }

        /// <summary>
        /// Abrir un fichero Word. Si nomFich no esta asignada crea un fichero nuevo y si
        /// no abre el que este asignado a la variable.
        /// </summary>
        public void AbrirFichero()
        {
            try
            {
                object missingValue = Type.Missing;

                // Abrir un fichero existente
                if (this.nomFich != null)
                {
                    object fichero = this.nomFich;
                    object visible = true;

#if OFFICEXP
					this.documento = this.word.Word.Documents.Open2000(ref fichero,
#else
                    this.documento = this.word.Word.Documents.Open(ref fichero,
#endif

                        ref missingValue,
                        ref missingValue, ref missingValue, ref missingValue,
                        ref missingValue, ref missingValue, ref missingValue,
                        ref missingValue, ref missingValue, ref missingValue,
                        ref visible, ref missingValue, ref missingValue, ref missingValue);

                    this.documentosWord = this.word.Word.Documents;
                    this.abierto = true;
                }
                else
                {
                    this.documento = this.word.Word.Documents.Add(ref missingValue, ref missingValue,
                        ref missingValue, ref missingValue);

                    this.documentosWord = this.word.Word.Documents;
                    this.abierto = true;
                }
            }
            catch (Exception ex)
            {
                this.abierto = false;
            }
        }

        #endregion

        #region Cambiar orientación de la página

        /// <summary>
        /// Cambia el tipo de horientacion a vertical
        /// </summary>
        public void CambiarOrientacionVertical()
        {
            this.CambiarOrientacion(WdOrientation.wdOrientLandscape);
        }

        /// <summary>
        /// Cambia el tipo de horientacion a horizontal
        /// </summary>
        public void CambiarOrientacionHorizontal()
                {
            this.CambiarOrientacion(WdOrientation.wdOrientPortrait);
            }

        /// <summary>
        /// Agranda el word para poder visualizarlo
        /// </summary>
        public void AgrandarAnchuraWord(int cols)
        {
            PageSetup objPagina = this.documentosWord[this.documentosWord.Count].PageSetup;
            objPagina.PageWidth = ((cols * 42 + objPagina.PageWidth) < 1584) ? cols * 42 + objPagina.PageWidth : 1584;
            this.documentosWord[this.documentosWord.Count].PageSetup = objPagina;
        }

        #endregion

        #region Procesado de Texto

        /// <summary>
        /// Busca las cadenas que se pasan y las sustituye
        /// </summary>
        /// <param name="cadenasBuscar"></param>
        /// <param name="cadenasReemplazar"></param>
        public void BuscaYSustituye(Dictionary<string, object> objRellenar, bool enDocumento, bool enFrames)
        {
            foreach (string auxValor in objRellenar.Keys)
            {
                BuscaYSustituye(auxValor, objRellenar[auxValor], enDocumento, enFrames);
            }
        }


        /// <summary>
        /// Busca y reemplaza la cadena a buscar por el texto re-emplazado, 
        /// </summary>
        /// <param name="cadenaBuscar"></param>
        /// <param name="cadenaReemplazar"></param>
        /// <param name="enDocumento"></param>
        /// <param name="enFrames"></param>
        public void BuscaYSustituye(object cadenaBuscar, object cadenaReemplazar, bool enDocumento, bool enFrames)
        {
            try
            {
                if (enDocumento)
                {
                    // Seleccionamos el contenido del documento
                    this.documentosWord[this.documentosWord.Count].Select();
                    Microsoft.Office.Interop.Word.Find fnd = this.word.Word.Selection.Find;
                    this.buscarYReemplazar(fnd, (string)cadenaBuscar, (string)cadenaReemplazar);
                    var prueba = this.documentosWord[1].Paragraphs.Last;
                    var prueba2 = prueba.Range.End;
                }

                // Frames agrupados y separados
                if (enFrames)
                {
                    foreach (Microsoft.Office.Interop.Word.Shape sh in this.documentosWord[this.documentosWord.Count].Shapes)
                    {
                        Microsoft.Office.Interop.Word.Find fnd;

                        try
                        {
                            // Seleccionamos el contenido del documento
                            this.documentosWord[this.documentosWord.Count].Select();
                            if (sh.GroupItems.Count > 0)
                            {
                                System.Collections.ArrayList ar = DesagruparShapes(sh.GroupItems, null);

                                foreach (Microsoft.Office.Interop.Word.Shape ssh in ar)
                                {
                                    if (ssh.TextFrame != null)
                                    {
                                        Microsoft.Office.Interop.Word.TextFrame f = ssh.TextFrame;
                                        Microsoft.Office.Interop.Word.Range r = f.ContainingRange;
                                        fnd = r.Find;

                                        this.buscarYReemplazar(fnd, (string)cadenaBuscar, (string)cadenaReemplazar);
                                    }
                                }
                            }
                        }
                        catch
                        {
                            try
                            {
                                if (sh.TextFrame != null)
                                {
                                    Microsoft.Office.Interop.Word.TextFrame f = sh.TextFrame;
                                    Microsoft.Office.Interop.Word.Range r = f.ContainingRange;
                                    fnd = r.Find;

                                    this.buscarYReemplazar(fnd, (string)cadenaBuscar, (string)cadenaReemplazar);
                                }
                            }
                            catch
                            {
                            }
                        }
                    }

                }

            }
            catch
            {

            }
        }
        
        #endregion

        #region Convertir a formatos

        /// <summary>
        /// Convierte el word actual en formato PDF en la ruta seleccionada
        /// </summary>
        /// <param name="rutaDestino"></param>
        public void ConvertirA_PDF(string rutaDestino)
        {
            this.ConvertirFormato(rutaDestino, WdExportFormat.wdExportFormatPDF);            
        }

        /// <summary>
        /// Convierte el word actual en formato XPS en la ruta seleccionada
        /// </summary>
        /// <param name="rutaDestino"></param>
        public void ConvertirA_XPS(string rutaDestino)
        {
            this.ConvertirFormato(rutaDestino, WdExportFormat.wdExportFormatXPS);
        }

        #endregion

        #region Insertar Contenido de un Documento en Otro


        /// <summary>
        /// Inserta el contenido de un documento en otro sustituyendo el texto que se pasa
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="texto"></param>
        /// <param name="formaSustitucion"></param>
        /// <returns></returns>
        public bool InsertarDocumentoSustituyendoTexto(string doc,
            string texto, int formaSustitucion)
        {
            try
            {
                // Seleccionamos el contenido del documento
                this.documentosWord[this.documentosWord.Count].Select();
                this.BuscarTextoEInsertarContenidoDocumento(
                    this.word.Word.Selection, texto, doc, formaSustitucion);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public int GetPagesNumber()
        {
            int result = 1;
            try
            {
                var activeDoc = this.documentosWord[this.documentosWord.Count];
                object missing = null;
                Range wholeDocRange = activeDoc.Range(ref missing, ref missing);
                result = (int)wholeDocRange.get_Information(WdInformation.wdNumberOfPagesInDocument);                
            }
            catch
            {
                result = 1;   
            }

            return result;
        }

        /// <summary>
        /// Concatenamos todos los ficheros pasados por parámetro en uno solo
        /// </summary>
        /// <param name="ficherosDocx"></param>
        /// <returns></returns>
        public bool ConcatenarDocx(System.Collections.ObjectModel.Collection<string> ficherosDocx)
        {
            Boolean procesoCorrecto = true;
            PageSetup[] pageSetupList = new PageSetup[ficherosDocx.Count];
            Range[] headersRange = new Range[ficherosDocx.Count];
                        
            try
            {    
                var rutaFichero = String.Copy(this.nomFich);
                this.nomFich = null;
                                
                this.AbrirFichero();               
                Microsoft.Office.Interop.Word.Range rng = this.documento.Range();
                var oWord = new Microsoft.Office.Interop.Word.Application();
                oWord.Visible = false;

                this.documento.UpdateStylesOnOpen = false;                                
                object oStart = Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseStart;        

                // Para cada fichero, creamos un reader y copiamos las páginas usando el writer 
                for (int i = ficherosDocx.Count - 1; i >= 0; i--)
                {
                    string fichero = ficherosDocx[i];
                    rng.InsertFile(FileName: fichero);                                                                                      
                    var oDoc2 = this.documentosWord.Add(fichero);
                    Microsoft.Office.Interop.Word.Range oRange = oDoc2.Content;
                    pageSetupList[i] = oRange.PageSetup;
                    rng.Collapse(ref oStart);
                }
                   
                //Recorremos las secciones del documento y asignamos el pagesetup de los doc originales
                for (int i = 0; i < this.documento.Sections.Count; i++)
                {
                    var seccionDocumento = this.documento.Sections[i + 1].PageSetup;
                    seccionDocumento.Orientation = pageSetupList[i].Orientation;
                    seccionDocumento.BottomMargin = pageSetupList[i].BottomMargin;
                    seccionDocumento.LeftMargin = pageSetupList[i].LeftMargin;
                    seccionDocumento.RightMargin = pageSetupList[i].RightMargin;
                    seccionDocumento.TopMargin = pageSetupList[i].TopMargin;
                    seccionDocumento.PageWidth = pageSetupList[i].PageWidth;
                }

                //JCNS. LOGO 
                //No funcionaba ni con el antiguo logo, desplazaba el logo hacia abajo. con esto se corrige. no me gusta nada pero tengo que entregarlo

                if (ConfigurationManager.AppSettings["WordEncabezadoDesdeArriba"] != null)
                {
                    this.documento.PageSetup.HeaderDistance = oWord.Application.CentimetersToPoints(float.Parse(ConfigurationManager.AppSettings["WordEncabezadoDesdeArriba"].ToString(), CultureInfo.InvariantCulture));
                    this.documento.PageSetup.FooterDistance = oWord.Application.CentimetersToPoints(float.Parse(ConfigurationManager.AppSettings["WordPieDePAginaDesdeAbajo"].ToString(), CultureInfo.InvariantCulture));
                }
                else
                {
                    this.documento.PageSetup.HeaderDistance = oWord.Application.CentimetersToPoints(float.Parse("0.1", CultureInfo.InvariantCulture));
                    this.documento.PageSetup.FooterDistance = oWord.Application.CentimetersToPoints(float.Parse("0.4", CultureInfo.InvariantCulture));
                }



                // Cerramos los objetos writer y document BUENO
                this.GuardarComo(rutaFichero);
                this.CerrarWord();
                              
                foreach (string fichero in ficherosDocx)
                {
                    if (System.IO.File.Exists(fichero))
                    {
                        System.IO.File.Delete(fichero);
                    }
                }                
            }
            catch (Exception e)
            {
                procesoCorrecto = false;
            }

            return procesoCorrecto;           
        }

        #endregion

        #region EscribirTextoAlFinalDelDocumento

        /// <summary>
        /// Escribe el texto que se pasa al final del documento
        /// </summary>
        /// <param name="texto"></param>
        public void EscribirTextoAlFinalDelDocumento(string texto)
        {
            Microsoft.Office.Interop.Word.Range rng = this.documentosWord[this.documentosWord.Count].Range(ref this.m_objOpt, ref this.m_objOpt);


            // Nos posicionamos al final del documento
            rng.SetRange(rng.End, rng.End);
            rng.Text = texto;
        }

        #endregion

        #region Insertar Saltos de página

        /// <summary>
        /// Escribe un salto de página al final del documento
        /// </summary>
        public void InsertarSaltoPaginaAlFinalDocumento()
        {
            Microsoft.Office.Interop.Word.Range rng = this.documentosWord[this.documentosWord.Count].Range(ref this.m_objOpt, ref this.m_objOpt);

            // Nos posicionamos al final del documento
            rng.SetRange(rng.End, rng.End);
            rng.Words.Last.InsertBreak(Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak);
        }

        /// <summary>
        /// Escribe un cambio de sección al final del documento
        /// </summary>
        public void InsertarCambioSeccionAlFinalDocumento()
        {
            Microsoft.Office.Interop.Word.Range rng = this.documentosWord[this.documentosWord.Count].Range(ref this.m_objOpt, ref this.m_objOpt);

            // Nos posicionamos al final del documento
            rng.SetRange(rng.End, rng.End);
            //rng.Words.Last.InsertBreak(Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage);
            rng.Words.Last.InsertBreak(Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage);
        }

        /// <summary>
        /// Insertar un salto de línea solo si no esta el cursor en una página ya vacia.
        /// </summary>
        public void InsertarSaltoPaginaAlFinalDocumentoSiNecesario()
        {
            object what = WdGoToItem.wdGoToPage;
            object which = WdGoToDirection.wdGoToFirst;
            object count = this.documentosWord[this.documentosWord.Count].ActiveWindow.Panes[1].Pages.Count;

            Range startRange = word.Word.Selection.GoTo(ref what, ref which, ref count, ref m_objOpt);
            object count2 = (int)count + 1;
            Range endRange = word.Word.Selection.GoTo(ref what, ref which, ref count2, ref m_objOpt);
            
            if (endRange.Start == startRange.Start)
            {
                which = WdGoToDirection.wdGoToLast;
                what = WdGoToItem.wdGoToLine;
                endRange = word.Word.Selection.GoTo(ref what, ref which, ref count2, ref m_objOpt);                
            }

            endRange.SetRange(startRange.Start, endRange.End); 
            endRange.Select();

            if (endRange.Words.Count > 0) 
            {
                Microsoft.Office.Interop.Word.Range rng = this.documentosWord[this.documentosWord.Count].Range(ref this.m_objOpt, ref this.m_objOpt);

                // Nos posicionamos al final del documento
                rng.SetRange(rng.End, rng.End);
                rng.Words.Last.InsertBreak(Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak);

                // Si nos ha insertado dos páginas con el salto de página
                if ((count as int?) + 2 == this.documentosWord[this.documentosWord.Count].ActiveWindow.Panes[1].Pages.Count)
                {
                    this.documentosWord[this.documentosWord.Count].Undo(1);
                }
            }
            
        }

        #endregion

        #region Manejo de Tablas

        /// <summary>
        /// Agrega en la primera fila de la tabla seleccionada sin cambiar el estilo de la tabla
        /// </summary>
        /// <param name="datos"></param>
        public void AgregarTituloSinFormatoATabla(string[] datos)
        {
            int i = 1;
            // La rellenamos
            foreach (string dato in datos)
            {
                Range obj = this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].Cell(0, i).Range;
                //System.Threading.Thread.Sleep(250);
                obj.Text = dato;
                i++;
            }
        }

        /// <summary>
        /// Devuelve el numero de columnas de la tabla seleccionada
        /// </summary>
        /// <returns></returns>
        public int ObtenerNumeroColumnasTabla()
        {
            return this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].Columns.Count;
        }

        /// <summary>
        /// Inserta una columna a la derecha de la columna que se indica en el parámetro.
        /// </summary>
        /// <param name="indice"></param>
        public void InsertarColumnaTabla(int indice)
        {
            this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].Columns.Add(this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].Columns[indice]);
        }

        public void QuitarNegritaColumnaTabla()
        {
            for (int i = 0; i < this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].Rows.Count; i++)
            {
                this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].Rows[i].Range.Font.Bold = 0;                  
            }                      
        }

        /// <summary>
        /// Selecciona una tabla existente del fichero word abierto
        /// </summary>
        /// <param name="indTabla"></param>
        public void SeleccionarTabla(int indTabla)
        {
            try
            {
                this.numTablaSeleccionada = indTabla;
                //this.documentosWord[this.numDocumentoSeleccionado].Tables[numTablaSeleccionada].Select();                
            }
            catch
            {

            }
        }

        /// <summary>
        /// Elimina la tabla del documento
        /// </summary>
        /// <param name="indTabla"></param>
        public void EliminarTabla(int indTabla)
        {
            try
            {
                this.documentosWord[this.documentosWord.Count].Tables[indTabla].Delete();
            }
            catch
            {

            }
        }

        /// <summary>
        /// Agrega la primera linea
        /// </summary>
        /// <param name="datos"></param>
        public void AgregarPrimeraFilaSinFormatoATabla(string[] datos)
        {
            // Si la tabla no esta asignada no hacemos nada
            if (this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada] == null)
                return;

            // Si el numero de datos no se corresponde con las columnas de la
            // tabla volvemos sin hacer nada
            if (this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].Columns.Count != datos.Length)
                return;

            // La rellenamos
            int i = 1;
            foreach (string dato in datos)
            {
                this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].Cell(2, i).Range.Text = dato;
                i++;
                //System.Threading.Thread.Sleep(250);
            }

        }

        /// <summary>
        /// Agrega una fila a la tabla seleccionada sin cambiar el estilo de la tabla
        /// </summary>
        /// <param name="datos"></param>
        public void AgregarFilaSinFormatoATabla(string[] datos)
        {
            // Si la tabla no esta asignada no hacemos nada
            if (this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada] == null)
                return;

            // Si el numero de datos no se corresponde con las columnas de la
            // tabla volvemos sin hacer nada
            if (this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].Columns.Count != datos.Length)
                return;

            // Creamos la fila
            int i = 1;
            Microsoft.Office.Interop.Word.Row laFila = this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].Rows.Add(ref this.m_objOpt);

            // La rellenamos
            foreach (string dato in datos)
            {
                this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].Cell(laFila.Index, i).Range.Text = dato;
                i++;
            }
        }

        /// <summary>
        /// Agranda la primera columna de la tabla para ajustarla al contenido.
        /// </summary>
        public void AgrandarColumnaTabla()
        {
            this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].Columns[1].SetWidth(100, WdRulerStyle.wdAdjustNone);
        }

        public void PonerTablaCabeceraRepetidaPublicorreo()
        {
            this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].Select();
            this.documentosWord[this.documentosWord.Count].ActiveWindow.Selection.Font.Bold = 0;

            Shading objValor = this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].Columns[1].Shading;


            this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].set_Style("Grid Table 2 - Accent 11");
            this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].Borders.InsideColor = WdColor.wdColorWhite;
            this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].Borders.OutsideColor = WdColor.wdColorWhite;
            this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].Rows[1].HeadingFormat = -1;
            this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].Rows[1].Select();
            this.documentosWord[this.documentosWord.Count].ActiveWindow.Selection.Font.Bold = 1;
            
            this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].Columns[1].Select();
            this.documentosWord[this.documentosWord.Count].ActiveWindow.Selection.Font.Bold = 1;

            this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].Columns[5].Select();
            this.documentosWord[this.documentosWord.Count].ActiveWindow.Selection.Font.Bold = 1;

            this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].Columns[9].Select();
            this.documentosWord[this.documentosWord.Count].ActiveWindow.Selection.Font.Bold = 1;

            this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].Columns[13].Select();
            this.documentosWord[this.documentosWord.Count].ActiveWindow.Selection.Font.Bold = 1;

            //this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].set_Style("Grid Table 2 - Accent 11");
            //this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].Borders.InsideColor = WdColor.wdColorWhite;
            //this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].Borders.OutsideColor = WdColor.wdColorWhite;
            //this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].Rows[1].HeadingFormat = -1;
            //this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].Rows[1].Select();
            //this.documentosWord[this.documentosWord.Count].ActiveWindow.Selection.Font.Bold = 1;

            //this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].Columns[17].Shading.BackgroundPatternColorIndex = objValor.BackgroundPatternColorIndex;
            //this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].Columns[17].Select();
            //this.documentosWord[this.documentosWord.Count].ActiveWindow.Selection.Font.Bold = 1;
        }


        //Se encarga de poner la primera fila de la tabla en formato de header repetida
        public void PonerTablaCabeceraRepetida()
        {
            this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].set_Style("Grid Table 2 - Accent 11");
            this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].Rows[1].HeadingFormat = -1;            
        }

        public void DarFormatoTablaGrupoDestino()
        {
            //Centramos la alinación de la segunda y la tercera filas
            var oTable = this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada];                        
            DarFormatoTablaPrecioCierto();
            oTable.Rows[2].Cells[1].VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalBottom;

            //Si es grupo de tramo con descuento sobre tarifa, incormporamos un salto de linea entre los literales grupo de tramo y descuento por destino
            if (oTable.Rows[2].Cells[1].Range.Text.Contains("$SALTO"))
            {
                oTable.Rows[2].Cells[1].Range.Text = oTable.Rows[2].Cells[1].Range.Text.Replace("$SALTO", "\n").Replace("\r", String.Empty);

                //Ponemos en negrita el literal "Grupo de Tramos
                for (int i = 1; i <= 16; i++)
                {
                    oTable.Rows[2].Cells[1].Range.Characters[i].Bold = 1;
                }

                //Quitamos negrita al literal Descuento aplicado sobre tarifa
                for (int i = 17; i <= oTable.Rows[2].Cells[1].Range.Text.Length-1; i++)
                {
                    oTable.Rows[2].Cells[1].Range.Characters[i].Bold = 0;
                }
            }

        }

        /// <summary>
        /// Configura un borde especifico pasado por parámetro
        /// </summary>        
        public void SetLineBorder(Table oTable, int iRow, int iCol, WdBorderType? borderType, WdColor borderColor)
        {
            var celda = oTable.Rows[iRow].Cells[iCol];

            if (borderType.HasValue)
            {
                celda.Borders[borderType.Value].LineStyle = WdLineStyle.wdLineStyleSingle;
                celda.Borders[borderType.Value].LineWidth = WdLineWidth.wdLineWidth025pt;
                celda.Borders[borderType.Value].Color = borderColor;
            }
            else
            {
                celda.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                celda.Borders.OutsideLineWidth = WdLineWidth.wdLineWidth025pt;
                celda.Borders.OutsideColor = borderColor;
            }
        }


        /// <summary>
        /// Da formato al descuento aplicado sobre tarifa para las tablas de descuento por destinos
        /// </summary>        
        public void DarFormatoTablaVA()
        {
            //Centramos la alinación de la segunda y la tercera filas
            var oTable = this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada];
            var oTableSel = this.word.Word.Selection;
            var numCols = oTable.Columns.Count;
            const WdColor COLOR_BORDE_EXTERIOR = WdColor.wdColorWhite;

            oTable.Borders.OutsideColor = COLOR_BORDE_EXTERIOR;
            oTable.Borders.InsideColor = COLOR_BORDE_EXTERIOR;
            oTable.Rows[1].Borders.InsideLineStyle = WdLineStyle.wdLineStyleNone;
            oTable.Rows[1].Borders.OutsideLineStyle = WdLineStyle.wdLineStyleNone;
            oTable.Rows[1].Range.Font.Color = WdColor.wdColorWhite;
            //JCNS. LOGO. Pongo el color corporativo
            //oTable.Rows[1].Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 69, 125));
            oTable.Rows[1].Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 36, 86));

            //Damos borde a los extremos de la tabla
            for (int i = 2; i <= oTable.Rows.Count; i++)
            {
                int numCells = oTable.Rows[i].Cells.Count;

                SetLineBorder(oTable, i, 1, WdBorderType.wdBorderLeft, COLOR_BORDE_EXTERIOR);
                SetLineBorder(oTable, i, numCells, WdBorderType.wdBorderRight, COLOR_BORDE_EXTERIOR);
                SetLineBorder(oTable, i, numCells, WdBorderType.wdBorderTop, COLOR_BORDE_EXTERIOR);
                SetLineBorder(oTable, i, numCells, WdBorderType.wdBorderBottom, COLOR_BORDE_EXTERIOR);
            }
            //DarSaltoCondicional();
            oTable.Rows[1].Cells[1].Range.Borders.InsideColor = WdColor.wdColorWhite;
            oTable.Rows[1].Cells[1].Range.Borders.OutsideColor = WdColor.wdColorWhite;
        }

        /// <summary>
        /// Da formato al descuento aplicado sobre tarifa para las tablas de descuento por destinos
        /// </summary>        
        public void DarFormatoTablaDescuentoDestino(bool esInformeTarifas, bool esPublicorreo=false)
        {          
            //Centramos la alinación de la segunda y la tercera filas
            var oTable = this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada];           
            var oTableSel = this.word.Word.Selection;
            var numCols = oTable.Columns.Count;
            const WdColor COLOR_BORDE_EXTERIOR = WdColor.wdColorWhite;

            oTable.Rows[2].Range.Font.Size = 6;
            oTable.Rows[2].Range.Font.Color = oTable.Rows[1].Range.Font.Color = COLOR_BORDE_EXTERIOR;
            oTable.Rows[2].Cells[1].Range.Font.Color = WdColor.wdColorBlack;
            oTable.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            oTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
            oTable.Borders.OutsideColor = COLOR_BORDE_EXTERIOR;
            oTable.Borders.InsideColor = COLOR_BORDE_EXTERIOR;
            
            oTable.Rows[1].Borders.InsideColor = COLOR_BORDE_EXTERIOR;
            oTable.Rows[2].Borders.OutsideColor = COLOR_BORDE_EXTERIOR;
            oTable.Rows[2].Borders.InsideColor = COLOR_BORDE_EXTERIOR;

            if (!esInformeTarifas)
            {
                oTable.Rows[3].Select();
                oTableSel.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                oTable.Rows[4].Select();
                oTableSel.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            }
            //Ponemos en negrita la cuarta fila
            oTableSel.Font.Bold = 1;
            
            //Cambiamos el background color de la primera y tercera filas         
            for (int i = 1; i <= oTable.Rows[1].Cells.Count; i ++)
            {
                Cell celda = (Cell) oTable.Rows[1].Cells[i];


                if (i > 1)
                {
                    try
                    {
                        celda.Borders.Enable = 1;

                        Borders borderUno = celda.Borders;
                        borderUno[WdBorderType.wdBorderBottom].LineStyle = borderUno[WdBorderType.wdBorderTop].LineStyle =
                        borderUno[WdBorderType.wdBorderLeft].LineStyle = borderUno[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;
                        borderUno[WdBorderType.wdBorderLeft].Color = borderUno[WdBorderType.wdBorderRight].Color = WdColor.wdColorWhite;
                        //border[WdBorderType.wdBorderBottom].LineStyle = border[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleNone;

                        //JCNS. LOGO. Pongo el color corporativo
                        //celda.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 69, 125));
                        celda.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 36, 86));
                        borderUno[WdBorderType.wdBorderLeft].LineWidth = borderUno[WdBorderType.wdBorderRight].LineWidth = WdLineWidth.wdLineWidth025pt;
                        
                    }
                    catch (Exception e)
                    {
                        System.Console.WriteLine("GENERADOR: Error al asignar borde en el informe. No afecta al funcionamiento de la aplicación.");
                    }

                    celda.Borders.OutsideColor =  COLOR_BORDE_EXTERIOR;

                    //esPublicorreo = true;
                    //if (esPublicorreo)
                    //{
                    celda = oTable.Rows[2].Cells[i];

                    //JCNS. LOGO. Pongo el color corporativo
                    //celda.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 69, 125));
                    celda.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 36, 86));
                    SetLineBorder(oTable, 2, i, null, COLOR_BORDE_EXTERIOR);                   
                    //}
                }              
            }

            //Nos aseguramos de que la primera celda tenga borde asignado
            if(oTable.Rows[1].Cells.Count > 2)
                oTable.Rows[1].Cells[2].Borders = oTable.Rows[1].Cells[3].Borders;


            
            //Mergeamos la cuarta fila (Controlamos que no se intente mergear la misma celda)
            if (!esInformeTarifas)
            {
                if (numCols != 2)
                    oTable.Rows[4].Cells[2].Merge(oTable.Rows[4].Cells[numCols]);
            }

            //Configuramos las cabeceras
            SetLineBorder(oTable, 1, 2, WdBorderType.wdBorderLeft, COLOR_BORDE_EXTERIOR); 
                        
            //Configuramos la alineación de la celda precio
            bool celdaConTextoPrecio = !(oTable.Rows[2].Cells[1].Range.Text == "Precio\r\a");
            if (celdaConTextoPrecio)
            {
                oTable.Rows[2].Cells[1].VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalBottom;
                oTable.Rows[2].Cells[1].Range.Font.Size = 9;
            }

            //Borramos bordes de celdas en blanco
            oTable.Rows[2].Cells[1].Shading.BackgroundPatternColor = WdColor.wdColorWhite;
            oTable.Rows[2].Cells[1].Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;

            //anyadirValTablaGT(objDestino.CodDestinoSAPSinZona);

            //Damos borde a los extremos de la tabla
            for (int i = 3; i <= oTable.Rows.Count; i++)
            {
                int numCells = oTable.Rows[i].Cells.Count;

                SetLineBorder(oTable, i, 1, WdBorderType.wdBorderLeft, COLOR_BORDE_EXTERIOR);
                SetLineBorder(oTable, i, numCells, WdBorderType.wdBorderRight, COLOR_BORDE_EXTERIOR);                 
            }

            //Si es publicorreo, quitamos el borde de la primera linea
            //if (esPublicorreo)
            //{
                //oTable.Rows[1].Cells[1].Borders[WdBorderType.wdBorderTop].Visible = false;
                //oTable.Rows[1].Cells[1].Borders[WdBorderType.wdBorderLeft].Visible = false;

                //oTable.Rows[4].Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(System.Drawing.Color.FromArgb(185, 207, 230));
                //oTable.Rows[4].Borders.InsideColor = COLOR_BORDE_EXTERIOR;
            //}
            //oTable.Rows[1].Cells[1].Range.Borders.InsideLineStyle = WdLineStyle.wdLineStyleNone;
            //oTable.Rows[1].Cells[1].Range.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleNone;
            
            oTable.Rows[1].Cells[1].Range.Borders.InsideColor = WdColor.wdColorWhite;
            oTable.Rows[1].Cells[1].Range.Borders.OutsideColor = WdColor.wdColorWhite;

            if (!esInformeTarifas)
            {
                //JCNS. LOGO. Pongo el color corporativo
                //oTable.Rows[3].Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 69, 125));
                oTable.Rows[3].Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 36, 86));
                oTable.Rows[3].Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
                oTable.Rows[3].Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                oTable.Rows[3].Borders.InsideColor = COLOR_BORDE_EXTERIOR;
                oTable.Rows[3].Borders.OutsideColor = COLOR_BORDE_EXTERIOR;
                oTable.Rows[3].Range.Font.Color = COLOR_BORDE_EXTERIOR;
                //JCNS. LOGO. Pongo el color corporativo
                //oTable.Rows[4].Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 69, 125));
                oTable.Rows[4].Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 36, 86));
                oTable.Rows[4].Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
                oTable.Rows[4].Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                oTable.Rows[4].Borders.InsideColor = COLOR_BORDE_EXTERIOR;
                oTable.Rows[4].Borders.OutsideColor = COLOR_BORDE_EXTERIOR;
                oTable.Rows[4].Range.Font.Color = COLOR_BORDE_EXTERIOR;
            }
            
        }

        /// <summary> ,
        /// Da formato al descuento aplicado sobre tarifa para las tablas de descuento por destinos
        /// </summary>        
        public void DarFormatoTablaPrecioCierto(bool esPublicorreo = false)
        {

            //Centramos la alinación de la segunda y la tercera filas
            var oTable = this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada];
            var oTableSel = this.word.Word.Selection;
            var numCols = oTable.Columns.Count;
            const WdColor COLOR_BORDE_EXTERIOR = WdColor.wdColorWhite;

            oTable.Rows[2].Range.Font.Size = 6;
            oTable.Rows[2].Range.Font.Color = COLOR_BORDE_EXTERIOR;
            oTable.Rows[2].Range.Font.Color = oTable.Rows[1].Range.Font.Color = COLOR_BORDE_EXTERIOR;
           
            oTable.Borders.OutsideColor = COLOR_BORDE_EXTERIOR;
            oTable.Borders.InsideColor = COLOR_BORDE_EXTERIOR;
            
            //Ponemos en negrita la cuarta fila
            oTableSel.Font.Bold = 1;

            for (int i = 1; i <= oTable.Rows[1].Cells.Count; i++)
            {
                Cell celda = (Cell)oTable.Rows[1].Cells[i];

                if (i > 1)
                {
                    try
                    {
                        celda.Borders.Enable = 1;

                        Borders borderUno = celda.Borders;
                        borderUno[WdBorderType.wdBorderBottom].LineStyle = borderUno[WdBorderType.wdBorderTop].LineStyle =
                        borderUno[WdBorderType.wdBorderLeft].LineStyle = borderUno[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;
                        borderUno[WdBorderType.wdBorderLeft].Color = borderUno[WdBorderType.wdBorderRight].Color = WdColor.wdColorWhite;
                        //border[WdBorderType.wdBorderBottom].LineStyle = border[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleNone;
                        //JCNS. LOGO. Pongo el color corporativo
                        //celda.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0,83,141));
                        celda.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 36, 86));
                        borderUno[WdBorderType.wdBorderLeft].LineWidth = borderUno[WdBorderType.wdBorderRight].LineWidth = WdLineWidth.wdLineWidth025pt;

                    }
                    catch (Exception e)
                    {
                        System.Console.WriteLine("GENERADOR: Error al asignar borde en el informe. No afecta al funcionamiento de la aplicación.");
                    }

                    celda.Borders.OutsideColor = COLOR_BORDE_EXTERIOR;

                    //esPublicorreo = true;
                    //if (esPublicorreo)
                    //{
                    celda = oTable.Rows[2].Cells[i];
                    //JCNS. LOGO. Pongo el color corporativo
                    //celda.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0,83,141));
                    celda.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 36, 86));
                    SetLineBorder(oTable, 2, i, null, COLOR_BORDE_EXTERIOR);
                    //}
                }             
            }

            
            //oTable.Rows[1].Cells[2].Borders = oTable.Rows[1].Cells[3].Borders;

            ////Borramos los bordes exteriores    
            bool celdaConTextoPrecio = !(oTable.Rows[2].Cells[1].Range.Text == "Precio\r\a");

            oTable.Rows[2].Cells[1].Shading.BackgroundPatternColor = WdColor.wdColorWhite;
            oTable.Rows[2].Cells[1].Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
         
            if (celdaConTextoPrecio)
            {
                oTable.Rows[2].Cells[1].VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalBottom;
                oTable.Rows[2].Cells[1].Range.Font.Size = 7;
                oTable.Rows[2].Range.Font.Color = COLOR_BORDE_EXTERIOR;
            }
         
            //Damos borde a los extremos de la tabla
            for (int i = 3; i <= oTable.Rows.Count; i++)
            {
                int numCells = oTable.Rows[i].Cells.Count;
                SetLineBorder(oTable, i, 1, WdBorderType.wdBorderLeft, COLOR_BORDE_EXTERIOR);
                SetLineBorder(oTable, i, numCells, WdBorderType.wdBorderRight, COLOR_BORDE_EXTERIOR);
            }

            //Si es publicorreo, quitamos el borde de la primera linea
           // if (esPublicorreo)
           // {
                Cell primeraCelda = oTable.Rows[1].Cells[1];
                Cell segundaCelda = oTable.Rows[1].Cells[2];
                           
                bool celdaConBordes = !(primeraCelda.Range.Text == "\r\a");
                //primeraCelda.Borders[WdBorderType.wdBorderTop].Visible = celdaConBordes;
                //primeraCelda.Borders[WdBorderType.wdBorderLeft].Visible = celdaConBordes;

                if (celdaConBordes)
                {
                    SetLineBorder(oTable, 1, 1, WdBorderType.wdBorderTop, COLOR_BORDE_EXTERIOR);
                    SetLineBorder(oTable, 1, 1, WdBorderType.wdBorderLeft, COLOR_BORDE_EXTERIOR);
                }
                else
                {
                    Borders borderprimeraCelda = primeraCelda.Borders;
                    borderprimeraCelda[WdBorderType.wdBorderBottom].LineStyle = borderprimeraCelda[WdBorderType.wdBorderTop].LineStyle =
                    borderprimeraCelda[WdBorderType.wdBorderLeft].LineStyle = borderprimeraCelda[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;
                    borderprimeraCelda[WdBorderType.wdBorderLeft].Color = borderprimeraCelda[WdBorderType.wdBorderRight].Color = WdColor.wdColorWhite;
                    //celda.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(System.Drawing.Color.FromArgb(185, 207, 230));
                    borderprimeraCelda[WdBorderType.wdBorderLeft].LineWidth = borderprimeraCelda[WdBorderType.wdBorderRight].LineWidth = WdLineWidth.wdLineWidth025pt;
                }
                                
                celdaConBordes = !(segundaCelda.Range.Text == "\r\a");

                if (celdaConBordes)
                {
                    SetLineBorder(oTable, 1, 2, WdBorderType.wdBorderTop, COLOR_BORDE_EXTERIOR);
                    SetLineBorder(oTable, 1, 2, WdBorderType.wdBorderLeft, COLOR_BORDE_EXTERIOR);
                }
                else
                {
                    Borders bordersegundaCelda = segundaCelda.Borders;
                    bordersegundaCelda[WdBorderType.wdBorderBottom].LineStyle = bordersegundaCelda[WdBorderType.wdBorderTop].LineStyle =
                    bordersegundaCelda[WdBorderType.wdBorderLeft].LineStyle = bordersegundaCelda[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;
                    bordersegundaCelda[WdBorderType.wdBorderLeft].Color = bordersegundaCelda[WdBorderType.wdBorderRight].Color = WdColor.wdColorWhite;
                    //celda.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(System.Drawing.Color.FromArgb(185, 207, 230));
                    bordersegundaCelda[WdBorderType.wdBorderLeft].LineWidth = bordersegundaCelda[WdBorderType.wdBorderRight].LineWidth = WdLineWidth.wdLineWidth025pt;
                }
                
                //Si es cabecera de columna, subimos el tamaño de la fuente 
                if (oTable.Rows[2].Cells[1].Range.Text == "Peso\r\a")
                {                 
                    oTable.Rows[2].Cells[1].Range.Font.Size = 9;
                }

                oTable.Rows[2].Cells[1].Range.Font.Color = WdColor.wdColorBlack;

                //segundaCelda.Borders[WdBorderType.wdBorderTop].Visible = celdaConBordes;
                //segundaCelda.Borders[WdBorderType.wdBorderLeft].Visible = celdaConBordes;

                //if (primeraCelda.Range.Text == String.Empty)
                //{
                //    oTable.Rows[1].Cells[1].Borders[WdBorderType.wdBorderTop].Visible = false;
                //    oTable.Rows[1].Cells[1].Borders[WdBorderType.wdBorderLeft].Visible = false;
                //}
           // }

                oTable.Rows[1].Cells[1].Range.Borders.InsideColor = WdColor.wdColorWhite;
                oTable.Rows[1].Cells[1].Range.Borders.OutsideColor = WdColor.wdColorWhite;
        }
        
        public void InsertarTablaParteSuperior(int posicion)
        {
            Range rango = this.documentosWord[this.documentosWord.Count].Tables[posicion].Range;                
            rango.SetRange(80, 80);
            this.documentosWord[this.documentosWord.Count].Tables.Add(rango, 1, 2, Type.Missing, Type.Missing);
        }

        public void CopiarTabla(int numTablaOriginal)
        {
            int rowsToGoDown = 2;
                        
            Range rango = this.documentosWord[this.documentosWord.Count].Tables[numTablaOriginal].Range;
            rango.Select();
            this.word.Word.Selection.Copy();
            //this.word.Word.Selection.TypeParagraph();
            //this.word.Word.Selection.Tables[numTablaOriginal].Select();
            

            Range Rng = this.word.Word.ActiveDocument.Characters.Last;
            Rng.Select();

            //oWord.Selection.TypeParagraph();
            
            this.word.Word.Selection.MoveDown(WdUnits.wdLine, rowsToGoDown);
            this.word.Word.Selection.Paste();
            this.word.Word.Selection.InsertAfter("\r");
            
        }

        /// <summary>
        /// Inserta la tabla pasada por parámetro en la posición dDeloie la tabla seleccionada actualmente
        /// </summary>
        /// <param name="tablaEnString">Tabla. Ej: Nomre\tLugar\tEdad\nAntonio\tNavarra\t36\nPepe\tBarcelona\t27</param>
        public void EscribirTablaDeString(string tablaEnString, bool diferenteFormatoParaPrimeraColumna, int? maxDecimales, bool esFormatoDesdeHasta = false)
        {
            // tablaEnString tiene sus columnas separadas por tabuladores y sus filas por newlines
            object tab = Microsoft.Office.Interop.Word.WdTableFieldSeparator.wdSeparateByTabs;
            Table tabla = this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada];
            PageSetup documento = this.documentosWord[this.documentosWord.Count].PageSetup;
                        
            Range rango = tabla.Range;
            Microsoft.Office.Interop.Word.Font fuente = rango.Font.Duplicate;
            fuente.Size = fuente.Size - 1;

            object formato = tabla.AutoFormatType;
            
            rango.SetRange(rango.End, rango.End);
            
            // Borramos la tabla para insertar la nuestra en la misma posición
            tabla.Delete();
           
            rango.Text = tablaEnString;
            object anchoCelda = 43;
            object verdadero = true;
            int numColumnas = 0;
            numColumnas = tablaEnString.Substring(0, tablaEnString.IndexOf('\n')).Split('\t').Length + 1;

            //Si el ancho del documento es el máximo posible, repartimos el tamaño del documento entre el número de columnas
            if (documento.PageWidth >= 1584)
            {   
                anchoCelda = 1584 / numColumnas; 
            }

            if (maxDecimales.HasValue && maxDecimales.Value > 3 && ((int)anchoCelda) < 43)
                fuente.Size = (maxDecimales.Value == 4) ? 6 : 5;
            
            rango.ConvertToTable(   ref tab,
                                    ref this.m_objOpt, ref this.m_objOpt, ref anchoCelda, ref formato,
                                    ref verdadero, ref this.m_objOpt, ref verdadero, ref this.m_objOpt,
                                    ref this.m_objOpt, ref this.m_objOpt, ref verdadero, ref this.m_objOpt,
                                    ref this.m_objOpt, ref this.m_objOpt, ref this.m_objOpt
                                );
            

            rango.Font = fuente;

            


            //[MMUNOZ] A partir de 60 columnas Word no puede controlar la tabla
            if (numColumnas <= 60)
            {
            // Alineamos todo a la derecha (La mayoría de las celdas son así)
            this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            }
            
            //[MMUNOZ] A partir de 60 columnas Word no puede controlar la tabla
            if (diferenteFormatoParaPrimeraColumna && numColumnas < 60)
            {
                // Ajustamos formato especial de la primera columna
                try
                {
                    this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].Columns[1].Width = 97;
                }
                catch (Exception e)
                {
                    this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].Columns[1].Width = 80;
                    System.Console.WriteLine(e.Message);                    
                }

                this.documentosWord[this.documentosWord.Count].Tables[this.numTablaSeleccionada].Columns[1].Select();
                this.word.Word.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

                //if (esFormatoDesdeHasta)
                //    this.word.Word.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                //else
                //    this.word.Word.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            }
            
        }

        #endregion

        #region Guardar el Libro



        /// <summary>
        /// Guarda el documento
        /// </summary>
        public void SalvarDocumento()
        {
            // Guardar el libro
            this.documentosWord[this.documentosWord.Count].Save();
        }

        #endregion

        #region Cerrar el Word

        /// <summary>
        /// Cierra la aplicacion Word
        /// </summary>
        public void CerrarWord()
        {
            object guardar = false;

            if (this.documentosWord[this.documentosWord.Count] != null)
                this.documentosWord[this.documentosWord.Count].Close(ref guardar, ref this.m_objOpt, ref this.m_objOpt);
            System.Threading.Thread.Sleep(500);

            this.word.Word.Quit(ref guardar, ref this.m_objOpt, ref this.m_objOpt);
            //System.Threading.Thread.Sleep(500);
        }

        #endregion

        #endregion

        #region Métodos Privados

        #region BuscarTextoEInsertarContenidoDocumento

        /// <summary>
        /// Busca en la seleccion la "cadenaBuscar" la borra e inserta el contenido entero del documento
        /// word que se pasa en el string documento
        /// </summary>
        /// <param name="sel"></param>
        /// <param name="cadenaBuscar"></param>
        /// <param name="documento"></param>
        protected void BuscarTextoEInsertarContenidoDocumento(Microsoft.Office.Interop.Word.Selection sel,
            string cadenaBuscar, string documento, int formaSustitucion)
        {
            //-----------------------------------------------------------------
            // Abrir el documento que se pasa y seleccionar todo el contenido
            //-----------------------------------------------------------------
            object missing = Type.Missing;
            object visible = true;
            object fichero = documento;

            //this.VerWord = true;

#if OFFICEXP	
				Word.Document doc = this.word.Word.Documents.Open2000(ref fichero,
#else
            Microsoft.Office.Interop.Word.Document doc = this.word.Word.Documents.Open(ref fichero,
#endif
 ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref visible, ref missing,
                ref missing, ref missing);

            // Seleccionar todo el documento y copiarlo al portapapeles
            object start = 0;
            object end = doc.Characters.Count - 1;
            doc.Range(ref start, ref end).Select();



            // Seleccionamos todo el texto
            this.word.Word.Selection.Copy();

            // Obtenemos texto, formato del párrafo y fuente
            string texto = this.word.Word.Selection.Text;
            Microsoft.Office.Interop.Word.ParagraphFormat pf = this.word.Word.Selection.ParagraphFormat;
            Microsoft.Office.Interop.Word.Font f = this.word.Word.Selection.Font;



            // Volvemos a seleccionar el documento original y cargamos el texto a buscar
            this.documentosWord[this.documentosWord.Count].Select();
            Microsoft.Office.Interop.Word.Find fnd = sel.Find;
            fnd.ClearFormatting();
            fnd.Text = cadenaBuscar;

            // Buscamos el texto y lo sustituimos por lo que teniamos copiado en el portapapeles
            while (fnd.Execute(ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing))
            {
                if (formaSustitucion == (int)FormasSustitucionTexto.CambiarTextoYEstilos)
                {
                    sel.Text = texto;
                    sel.ParagraphFormat = pf;
                    sel.Font = f;
                }
                else if (formaSustitucion == (int)FormasSustitucionTexto.CopiarPegarConFormato)
                {
                    sel.PasteAndFormat(Microsoft.Office.Interop.Word.WdRecoveryType.wdFormatOriginalFormatting);
                }
                else if (formaSustitucion == (int)FormasSustitucionTexto.CopiarSiguiendoLista)
                {
                    sel.PasteAndFormat(Microsoft.Office.Interop.Word.WdRecoveryType.wdListCombineWithExistingList);
                }
                else if (formaSustitucion == (int)FormasSustitucionTexto.CopiarPegarNormal)
                {
                    sel.Paste();
                }
                else if (formaSustitucion == (int)FormasSustitucionTexto.CambiarTextoMantenerEstilo)
                {
                    sel.Text = texto;
                }
            }

            doc.Close(ref missing, ref missing, ref missing);
        }

        #endregion

        #region Busquedas y Acciones


        /// <summary>
        /// Busca y reemplaza por el texto dado
        /// </summary>
        /// <param name="fnd"></param>
        /// <param name="cadenaBuscar"></param>
        /// <param name="cadenaReemplazar"></param>
        protected void buscarYReemplazar(Microsoft.Office.Interop.Word.Find fnd, string cadenaBuscar, string cadenaReemplazar)
        {
            Microsoft.Office.Interop.Word.Find fndOrigen = fnd;

            object replace = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;
            object missing = Type.Missing;

            if (cadenaReemplazar.Length <= 250)
            {
                fnd.ClearFormatting();
                fnd.Text = cadenaBuscar;
                fnd.Replacement.ClearFormatting();
                fnd.Replacement.Text = cadenaReemplazar;

                // Sustituimos
                while (fnd.Execute(ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref replace,
                    ref missing, ref missing, ref missing, ref missing))
                {
                }
            }
            else
            {
                // Debemos partir la cadena de reemplazo en trozos menores de 250 caracteres e ir sustituyendo
                int indice = 0;
                int tramo = 240;
                string morcillaBusqueda = "$$€€KKKK";
                string aux;
                bool seguir = true;
                int i = 0;
                int queda;

                while (seguir)
                {
                    fnd = fndOrigen;

                    // Obtenemos la cadena que queremos sustituir
                    queda = (cadenaReemplazar.Length - indice);
                    if (queda >= tramo)
                        aux = cadenaReemplazar.Substring(indice, tramo);
                    else
                        aux = cadenaReemplazar.Substring(indice, queda);

                    // Si hemos leido un tramo entero debemos ponerle la cadena que buscaremos luego
                    if (aux.Length == tramo)
                        aux += morcillaBusqueda + string.Format(CultureInfo.InvariantCulture, "{0:00}", i);
                    else
                        seguir = false;

                    // Rellenamos el campo de busqueda
                    fnd.ClearFormatting();

                    if (i == 0)
                        fnd.Text = cadenaBuscar;
                    else
                        fnd.Text = morcillaBusqueda + string.Format(CultureInfo.InvariantCulture, "{0:00}", i - 1);

                    fnd.Replacement.ClearFormatting();
                    fnd.Replacement.Text = aux;

                    // Sustituimos
                    while (fnd.Execute(ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref replace, ref missing, ref missing, ref missing, ref missing))
                    {
                    }

                    // Actualizamos el indice
                    indice += tramo;
                    i++;
                }
            }
        }        

        #endregion

        #region Cambiar orientación de la página

        /// <summary>
        /// cambia el tipo de orientacion
        /// </summary>
        /// <param name="tipoOrientacion"></param>
        private void CambiarOrientacion(WdOrientation tipoOrientacion)
        { 
            PageSetup objPagina = this.documentosWord[this.documentosWord.Count].PageSetup;
            objPagina.Orientation = tipoOrientacion;
            float ancho = objPagina.PageHeight;
            float alto = objPagina.PageWidth;
            objPagina.PageHeight = alto;
            objPagina.PageWidth = ancho;
            this.documentosWord[this.documentosWord.Count].PageSetup = objPagina;
        }

        #endregion

        #endregion

        #region DesagruparShapes

        /// <summary>
        /// Devuelve un array list con las shapes desagrupadas
        /// </summary>
        /// <param name="shapesAgrupadas"></param>
        /// <returns></returns>
        private System.Collections.ArrayList DesagruparShapes(Microsoft.Office.Interop.Word.GroupShapes shapesAgrupadas, System.Collections.ArrayList ar)
        {
            System.Collections.ArrayList a = ar;

            if (a == null)
                a = new System.Collections.ArrayList();

            foreach (Microsoft.Office.Interop.Word.Shape sh in shapesAgrupadas)
            {
                try
                {
                    if (sh.GroupItems.Count > 0)
                        DesagruparShapes(sh.GroupItems, a);
                }
                catch
                {
                    a.Add(sh);
                }
            }

            return a;
        }

        #endregion

        #region ConvertirFormato

        /// <summary>
        /// Convierte el documento actual en el formato y la ruta que se pasa por parámetro.
        /// </summary>
        /// <param name="rutaDestino"></param>
        /// <param name="formato"></param>
        private void ConvertirFormato(string rutaDestino, WdExportFormat formato)
        {
            this.documentosWord[this.documentosWord.Count].ExportAsFixedFormat(rutaDestino, formato);            
        }

        /// <summary>
        /// Guarda el documento actual en la ruta que se pasa por parámetro
        /// </summary>
        /// <param name="rutaDestino"></param>        
        public void GuardarComo(string rutaDestino)
        {
            this.documentosWord[this.documentosWord.Count].SaveAs2(rutaDestino);
        }

        #endregion
    }
}

