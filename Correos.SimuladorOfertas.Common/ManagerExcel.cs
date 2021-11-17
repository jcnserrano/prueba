using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace Correos.SimuladorOfertas.Common
{
    public class ManagerExcel
    {

        #region Estático
        
        public static void AbrirExcelStandalone(string ficheroExcel)
        {
            if (System.IO.File.Exists(ficheroExcel))
            {
                Process procesoExcel = new Process();
                // Usamos /x para lanzar el excel en una nueva instancia
                procesoExcel.StartInfo.Arguments = "/x \"" + ficheroExcel + "\"";
                procesoExcel.StartInfo.FileName = ManagerExcel.PathEjecutableExcel;
                procesoExcel.Start();
            }
        }

        /// <summary>
        /// Path del .exe de excel
        /// </summary>
        public static string PathEjecutableExcel
        {
            get
            {
                string pathExcel = string.Empty;
                Process[] procesosCorriendo = Process.GetProcesses();
                // Lo buscamos primero en los procesos corriendo
                foreach (Process proceso in procesosCorriendo)
                {
                    if (proceso.ProcessName == "EXCEL")
                    {
                        pathExcel = proceso.MainModule.FileName;
                        break;
                    }
                }

                // Si no está corriendo (Lo cual no debería pasar si está lanzado el simulador), se crea una instancia vacía y se mata
                if (pathExcel == string.Empty)
                {
                    Type tipoDeExcel = Type.GetTypeFromProgID("Excel.Application");
                    dynamic xlApp = Activator.CreateInstance(tipoDeExcel);
                    xlApp.Visible = false;
                    pathExcel = xlApp.Path + @"\Excel.exe";
                    xlApp.Quit();
                }
                return pathExcel;
            }
        }

        #endregion

        #region Propiedades

        // Aplicacion excel
        protected AplicacionExcel excel;
        // Libro excel
        public Microsoft.Office.Interop.Excel._Workbook excelLibro;
        // Hojas del libro excel
        public Microsoft.Office.Interop.Excel.Sheets excelHojas;
        // Hoja concreta del libro excel
        public Microsoft.Office.Interop.Excel.Worksheet excelHojaActual;        

        object m_objOpt;

        // Nombre del fichero
        protected string nomFich;

        // Indica si hay un fichero excel abierto
        protected bool abierto;

        /// <summary>
        /// Variable para ver la aplicación Word.
        /// </summary>
        public bool VerExcel
        {
            get { return this.excel.Excel.Visible; }
            set { this.excel.Excel.Visible = value; }
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

        /// <summary>
        /// Tamaño letra
        /// </summary>
        public int fontSize = 8;

        /// <summary>
        /// Anchura Columna
        /// </summary>
        public int columnWidth = 10;

        /// <summary>
        /// Rango de la celda con la que se trabaja
        /// </summary>
        Excel.Range rangeCelda;

        #region Tabla actual

        /// <summary>
        /// Información de si la tabla tiene cabecera
        /// </summary>
        bool UsarCabecera { get; set; }

        /// <summary>
        /// Información de si la tabla es numérica. Implica cambios de estilo
        /// </summary>
        bool TablaNumerica { get; set; }

        /// <summary>
        /// Número de columnas en la tabla actual
        /// </summary>
        int NumColumnas { get; set; }

        /// <summary>
        /// Indica cuántas columnas excel necesita el sting más largo de cada una de las columnas de la tabla
        /// </summary>
        int[] ColumnasRequeridas;

        public int[] Get_ColumnasRequeridas
        {
            get { return ColumnasRequeridas; }            
        }

        /// <summary>
        /// Define si se está dibujando una tabla
        /// </summary>
        bool EstaDibujandoTabla { get; set; }

        /// <summary>
        /// Ancho de una celda normal, sin hacer merge
        /// </summary>
        double AnchoCeldaBase { get; set; }

        /// <summary>
        /// Colección contenedora de las filas de la tabla. El array de objetos son las columnas
        /// </summary>
        System.Collections.ObjectModel.Collection<object[]> FilasTabla;

        /// <summary>
        /// Número de decimales a los que redondear
        /// </summary>
        int Redondeo { get; set; }

        /// <summary>
        /// Conjunto de columnas que tienen que estar en negrita
        /// </summary>
        Collection<int> ColumnasNegrita { get; set; }

        /// <summary>
        /// Guarda el tamaño en puntos que ocupa un dígito decimal de n cifras para no tener que calcularlo cada vez
        /// </summary>
        private Dictionary<int, double> TamOcupado { get; set; }

        /// <summary>
        /// Guarda la primera celda de una tabla finalizada
        /// </summary>
        private Excel.Range PrimeraCeldaTabla;

        public Excel.Range Get_PrimeraCeldaTabla
        {
            get { return PrimeraCeldaTabla; }          
        }

        #endregion

        #endregion

        #region Colores Diseño

        public int colorTextoCeldaEditable = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
        public int colorBordeCeldaEditable = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

        #endregion

        #region Constructores

        /// <summary>
        /// Construcor usado para generar nuevos excel desde plantillas
        /// </summary>
        /// <param name="nF"></param>
        public ManagerExcel(string nF)
        {

            this.nomFich = nF;
            // Crear Hoja Excel
            this.excel = new AplicacionExcel();
            this.excel.Excel.DisplayAlerts = false;
            this.excel.Excel.Visible = false;

            this.m_objOpt = System.Reflection.Missing.Value;

            TamOcupado = new Dictionary<int, double>();
        }

        /// <summary>
        /// Construcor usado para generar nuevos excel desde plantillas
        /// </summary>
        /// <param name="nF"></param>
        /// <param name="visible"></param>
        public ManagerExcel(string nF, bool visible)
        {

            this.nomFich = nF;
            // Crear Hoja Excel
            this.excel = new AplicacionExcel();
            this.excel.Excel.Visible = visible;
            this.excel.Excel.DisplayAlerts = visible;

            this.m_objOpt = System.Reflection.Missing.Value;

            TamOcupado = new Dictionary<int, double>();
        }

        /// <summary>
        /// Constructor usado para utilizar excel que ya han sido abiertos.
        /// </summary>
        /// <param name="excel"></param>
        public ManagerExcel(Microsoft.Office.Interop.Excel.Application excel)
        {
            // Crear Hoja Excel
            this.excel = new AplicacionExcel(excel);
            this.excel.Excel.DisplayAlerts = true;
            this.excelHojas = excel.Worksheets;
            this.m_objOpt = System.Reflection.Missing.Value;

            TamOcupado = new Dictionary<int, double>();
        }

        #endregion

        #region Métodos Públicos

        #region Apertura de Fichero Existente

        /// <summary>
        /// Abrir un fichero Excel
        /// </summary>
        public void AbrirFichero()
        {
            try
            {
                this.excelLibro = this.excel.Excel.Workbooks.Open(this.nomFich, 0,
                    false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true,
                    false, 0, true, false, false);

 

                this.excelHojas = this.excelLibro.Worksheets;

                this.abierto = true;
            }
            catch
            {
                this.abierto = false;
            }
        }

        #endregion

        #region Manejo de Hoja

        /// <summary>
        /// Seleccionar la primera hoja
        /// </summary>
        public void SeleccionarPrimeraHoja()
        {
            this.excelHojaActual = (Microsoft.Office.Interop.Excel.Worksheet)this.excelHojas.get_Item(1);
        }

        /// <summary>
        /// Selecciona la primera hoja si está libre. En caso contrario crea y selecciona una nueva hoja. Para usar desde informe múltiple.
        /// </summary>
        public void SeleccionarPrimeraHojaLibreOCrearNuevaYRenombrar(string nuevoNombre)
        {
            if (((Microsoft.Office.Interop.Excel.Worksheet)this.excelHojas.get_Item(1)).Name.StartsWith("Sheet"))
            {
                this.excelHojaActual = (Microsoft.Office.Interop.Excel.Worksheet)this.excelHojas.get_Item(1);
            }
            else
            {
                this.excelHojas.Add(Type.Missing, this.excelHojas.get_Item(this.excelHojas.Count), 1, Type.Missing);
                this.excelHojaActual = (Microsoft.Office.Interop.Excel.Worksheet)this.excelHojas.get_Item(this.excelHojas.Count);
                this.excelHojaActual.Cells.Interior.ColorIndex = 2; // Fondo blanco
            }

            this.excelHojaActual.Name = nuevoNombre;
        }

        /// <summary>
        /// Seleccionar la hoja que se pasa del documento
        /// </summary>
        /// <param name="hoja"></param>
        public void SeleccionarHoja(string hoja)
        {
            this.excelHojaActual = (Microsoft.Office.Interop.Excel.Worksheet)this.excelHojas.get_Item(hoja);
        }

        /// <summary>
        /// Desprotege la hoja con la contraseña que se pasa
        /// </summary>
        /// <param name="hoja"></param>
        /// <param name="password"></param>
        public void DesprotegerHoja(string password)
        {
            this.excelHojaActual.Unprotect(password);
        }
        
        /// <summary>
        /// protege la hoja con la contraseña que se pasa
        /// </summary>
        /// <param name="hoja"></param>
        /// <param name="password"></param>
        public void ProtegerHoja(string password)
        {
            this.excelHojaActual.Protect(password, true);
        }

        #endregion

        #region Procesado de Texto

        /// <summary>
        /// Busca las cadenas que se pasan y las sustituye
        /// </summary>
        /// <param name="cadenasBuscar"></param>
        /// <param name="cadenasReemplazar"></param>
        public void BuscaYSustituye(Dictionary<string, string> objRellenar, bool enDocumento, bool enFrames)
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
        public void BuscaYSustituye(string cadenaBuscar, string cadenaReemplazar, bool enDocumento, bool enFrames)
        {
            try
            {
                if (enDocumento)
                {
                    Excel.Range rangoCompleto = excelHojaActual.UsedRange;

                    Excel.Range celdaEncontrada = rangoCompleto.Find(cadenaBuscar, Type.Missing,
                        Type.Missing, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByColumns,
                        Excel.XlSearchDirection.xlNext, false, false, Type.Missing);

                    if (celdaEncontrada != null)
                    {
                        try 
                        {
                            celdaEncontrada.Value2 = (celdaEncontrada.Value2 as string).Replace(cadenaBuscar, cadenaReemplazar);
                        }
                        catch
                        {
                            celdaEncontrada.Value2 = cadenaReemplazar;
                        }
                    }
                }
            }
            catch
            {

            }
        }

        /// <summary>
        /// Escribe el texto pasado por parámetro en la siguiente file disponible
        /// </summary>
        /// <param name="p"></param>
        public void EscribirTextoAlFinalDelDocumento(string texto)
        {
            Excel.Range rango = excelHojaActual.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            // Movemos esa celda dos para abajo (para dejar margen de 1) y la posicionamos a la izquierda (con margen de 1)
            rango = excelHojaActual.Cells[rango.Row + 2, 2];
            rango.Value2 = texto;
        }

        /// <summary>
        /// Insertar el documento de la ruta en al final de la hoja actual
        /// </summary>
        /// <param name="rutaSegundoFichero"></param>
        /// <param name="sheetACopiar"></param>
        public void InsertarDocumentoAlFinal(string rutaSegundoFichero, int sheetACopiar = 1)
        {   
            Excel.Workbook workbookNuevo = this.excel.Excel.Workbooks.Open(rutaSegundoFichero);           
            Excel.Range rangoOrigen = workbookNuevo.Sheets[sheetACopiar].UsedRange;

            foreach (Excel.Shape o in workbookNuevo.Sheets[sheetACopiar].Shapes)
            {
                if (o.Type == Microsoft.Office.Core.MsoShapeType.msoPicture)
                    o.CopyPicture();

                excelHojaActual.Paste();                
            }

            // Buscamos la última celda usada
            Excel.Range rango = excelHojaActual.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);

            if (rango.Column == 1 && rango.Row == 1)
            {
                rango = excelHojaActual.Cells[rango.Row + 4, 2];
            }
            else
            {
                // Movemos esa celda dos para abajo (para dejar margen de 1) y la posicionamos a la izquierda (con margen de 1)
                rango = excelHojaActual.Cells[rango.Row + 2, 2];
            }

            rangoOrigen.Copy();
            rango.PasteSpecial();

            // Copiamos rango pequeño para evitar que excel dé el mensaje de que tienes copiada mucha información
            workbookNuevo.Sheets[sheetACopiar].Range["A1"].Copy();
            // Cerramos el workbook que acabamos de copiar
            workbookNuevo.Close();
        }

        /// <summary>
        /// Insertar el documento de la ruta en la línea en la que se encuentra el identificador
        /// </summary>
        /// <param name="rutaSegundoFichero"></param>
        /// <param name="identificador"></param>
        public void InsertarDocumentoEnFila(string rutaSegundoFichero, string identificador, int sheetACopiar = 1)
        {
            Excel.Worksheet worksheetId = null;
            Excel.Range rango = null;
            foreach (Excel.Worksheet worksheet in excelLibro.Worksheets)
            {
                foreach (Excel.Range r1 in worksheet.UsedRange)
                {
                    if (r1.Value2 == null)
                    {
                        continue;
                    }
                    if (r1.Value2.ToString().Equals(identificador))
                    {
                        rango = r1;
                        worksheetId = worksheet;
                        break;
                    }
                }
                if (rango != null)
                {
                    break;
                }
            }

            if (rango == null || worksheetId == null)
            {
                return;
            }

            Excel.Workbook workbookNuevo = this.excel.Excel.Workbooks.Open(rutaSegundoFichero);
            Excel.Range rangoOrigen = workbookNuevo.Sheets[sheetACopiar].UsedRange;

            rangoOrigen.Copy();
            rango.PasteSpecial();

            // Copiamos rango pequeño para evitar que excel dé el mensaje de que tienes copiada mucha información
            workbookNuevo.Sheets[sheetACopiar].Range["A1"].Copy();
            workbookNuevo.Close();
        }

        #endregion

        #region Manejo de Celdas

        /// <summary>
        /// Autoajusta todas las columnas de la hoja actual
        /// </summary>
        public void AutoAjustarColumnas()
        {
            this.excelHojaActual.Columns.AutoFit();
        }

        /// <summary>
        /// Escribir todo lo que viene el el diccionario de celdas con su valor protegiendo posteriormente la celda editada
        /// </summary>
        /// <param name="dCeldasValues"></param>
        public void EscribirCeldasProtegidasMultiple(Dictionary<string, object> objRellenar)
        {
            foreach (string auxValor in objRellenar.Keys)
            {
                EscribirEnCeldaProteger(auxValor, objRellenar[auxValor]);
            }
        }

        /// <summary>
        /// Escribir todo lo que viene el el diccionario de celdas con su valor
        /// </summary>
        /// <param name="dCeldasValues"></param>
        public void EscribirCeldasMultiple(Dictionary<string, object> objRellenar)
        {
            foreach (string auxValor in objRellenar.Keys)
            {
                EscribirEnCelda(auxValor, objRellenar[auxValor]);
            }
        }

        /// <summary>
        /// Escribir lo que se pasa en la celda que se pasa, protegiendo posteriormente la celda editada
        /// </summary>
        /// <param name="celda"></param>
        /// <param name="valor"></param>
        public void EscribirEnCelda(string celda, object valor)
        {
            // Obtenemos la celda
            Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)this.excelHojaActual.get_Range(celda);

            if (excelCell != null)
            {
                // Insertamos el valor en la celda
                excelCell.Value2 = valor;
            }
        }

        /// <summary>
        /// Se encarga de dibujar en la celda correspondiene el valor introducido como editable
        /// </summary>
        /// <param name="celdaInicial"></param>
        /// <param name="celdaFinal"></param>
        /// <param name="valor"></param>
        public void EscribirEnCeldaEditable(string celdaInicial, string celdaFinal, object valor)
        {
            EscribirEnCeldaEditableFormato(celdaInicial, celdaFinal, valor, "General");
        }

        /// <summary>
        /// Se encarga de dibujar en la celda correspondiene el valor introducido como editable
        /// </summary>
        /// <param name="celdaInicial"></param>
        /// <param name="celdaFinal"></param>
        /// <param name="valor"></param>
        public void EscribirEnCeldaEditableFormato(string celdaInicial, string celdaFinal, object valor, string formato)
        {
            rangeCelda = this.excelHojaActual.Cells.get_Range(celdaInicial, celdaFinal);

            if (celdaInicial != celdaFinal)
                rangeCelda.Merge();

            rangeCelda.Value = valor;
            rangeCelda.Font.Color = colorTextoCeldaEditable;
            rangeCelda.Borders.Color = colorBordeCeldaEditable;
            rangeCelda.Locked = false;
            rangeCelda.NumberFormat = formato;
        }

        /// <summary>
        /// Escribir lo que se pasa en la celda que se pasa, protegiendo posteriormente la celda editada
        /// </summary>
        /// <param name="celda"></param>
        /// <param name="valor"></param>
        public void EscribirEnCeldaProteger(string celda, object valor)
        {
            // Obtenemos la celda
            Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)this.excelHojaActual.get_Range(celda);

            if (excelCell != null)
            {
                // Insertamos el valor en la celda
                excelCell.Value2 = valor;

                // Se protege la celda
                excelCell.Locked = true;
            }
        }

        /// <summary>
        /// Escribir lo que se pasa en la combinación de celdas que se pasa, protegiendo posteriormente la celda editada
        /// </summary>
        /// <param name="celda"></param>
        /// <param name="valor"></param>
        public void EscribirEnCeldaProteger(string celdaInicial, string celdaFinal, object valor)
        {
            // Obtenemos la celda
            Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)this.excelHojaActual.get_Range(celdaInicial, celdaFinal);

            if (excelCell != null)
            {
                // Se combinan las celdas
                excelCell.Merge();

                // Insertamos el valor en la celda
                excelCell.Value2 = valor;

                // Se protege la celda
                excelCell.Locked = true;
            }
        }

        /// <summary>
        /// Lee la celda que se pasa
        /// </summary>
        /// <param name="celda"></param>
        /// <returns></returns>
        public string LeerDeCelda(string celda)
        {
            try
            {
                // Obtenemos la celda
                Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)this.excelHojaActual.get_Range(celda, celda);

                // Obtenemos el valor de la celda
                return excelCell.Value2.ToString();
            }
            catch
            {
                return "";
            }
        }

        /// <summary>
        /// Método que protege contra escritura la celda pasada por parámetro
        /// </summary>
        /// <param name="celda"></param>
        public void ProtegerCelda(string celda)
        {
            // Obtenemos la celda
            Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)this.excelHojaActual.get_Range(celda);

            if (excelCell != null)
            {
                excelCell.Locked = true;
            }
        }

        #endregion

        #region Tablas

        /// <summary>
        /// Comienza el proceso de dibujar una tabla
        /// </summary>
        public void IniciarDibujarTabla(bool estiloCabecera = false, bool tablaNumerica = false, int redondeo = 5)
        {
            NumColumnas = -1;
            EstaDibujandoTabla = true;
            ColumnasRequeridas = null;
            AnchoCeldaBase = excelHojaActual.Range["A1"].ColumnWidth;
            FilasTabla = new System.Collections.ObjectModel.Collection<object[]>();
            UsarCabecera = estiloCabecera;
            TablaNumerica = tablaNumerica;
            Redondeo = redondeo;
            ColumnasNegrita = new Collection<int>();
        }

        /// <summary>
        /// Devuelve la columna teniendo en cuenta los merges
        /// </summary>
        /// <param name="columnaCelda">Número de la columna sin tener en cuenta merges</param>
        /// <returns></returns>
        int ColumnaReal(int columnaCelda)
        {
            int columnaReal = 0;
            for (int i = 0; i < columnaCelda; i++)
            {
                columnaReal += ColumnasRequeridas[i];
            }
            return columnaReal;
        }
        
        private void AplicarEstilo(Excel.Range rango, int fila, int columna, Boolean EsTablaGrupoTramo = false, Boolean EsPublicorreo = true)
        {
            //JCNS. LOGO. Pongo el color corporativo
            //var colorFondoAzul = System.Drawing.Color.FromArgb(0, 69, 125);
            var colorFondoAzul = System.Drawing.Color.FromArgb(0, 36, 86);
            var colorBordeAzul = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(149, 179, 215));
            var colorLetraBlanco = System.Drawing.Color.FromArgb(255, 255, 255);

            if (columna == 0 || ColumnasNegrita.Contains(columna))
            {
                rango.Font.Bold = true;

                if (EsTablaGrupoTramo && columna == 0 && !EsPublicorreo)
                {
                    rango.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                }

            }
            if (UsarCabecera && fila == 0)
            {
                if (columna != 0) // No centramos la primera columna
                {
                    rango.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                }
                rango.Font.Bold = true;
            }
            else if (fila != 0 || !UsarCabecera)
            {
                rango.Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 255));
                System.Drawing.Color colorFondo = System.Drawing.Color.White;
                if (!UsarCabecera)
                {
                    if (fila % 2 == 0)
                    {
                        colorFondo = System.Drawing.Color.FromArgb(219, 229, 241);
                    }
                }
                else
                {
                    if ((fila + 1) % 2 == 0)
                        colorFondo = System.Drawing.Color.FromArgb(219, 229, 241);
                }
                rango.Interior.Color = System.Drawing.ColorTranslator.ToOle(colorFondo);

                if (EsTablaGrupoTramo && columna == 0)
                {
                    rango.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                }
            }
            if (TablaNumerica && fila != 0)
            {
                if (columna == 0)
                {                    
                    if (EsPublicorreo && !EsTablaGrupoTramo)
                    {
                        rango.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    }
                    else
                    {
                        rango.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;   
                    }
                }
                else
                {                    
                    rango.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                }
            }
        }

        /// <summary>
        /// Aumenta el nº de celdas de la tabla grupos tramos para que coincida con la tabla de valores inferior pintada posteriormente
        /// </summary>
        /// <param name="RangoPrimeraCelda">Primera celda de la tabla</param>
        /// <param name="numFilasTabla">Nº filas que forman la tabla</param>
        /// <param name="numColumnasNuevas">Nº de columnas vacías a añadir</param>
        /// <param name="numColumnasActuales">Nº de columas que forman la tabla</param>
        public void AumentarTamanyoTablaGrupoTramos(Excel.Range RangoPrimeraCelda, int numFilasTabla, int numColumnasNuevas, int numColumnasActuales)
        {            
            for (int i = 1; i < numFilasTabla + 1; i++)
            {
                //Hacemos primero un shift de tantas celdas como núm de columnas nuevas
                Excel.Range CeldaAnterior = excelHojaActual.Cells[RangoPrimeraCelda.Row + i, 1];
                Excel.Range CeldaColumna = excelHojaActual.Cells[RangoPrimeraCelda.Row + i, 2];
                                              
                for (int j = 0; j < numColumnasNuevas; j++)
                    CeldaAnterior.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
                
                //Hacemos merge de las celdas antiguas con las nuevas añadidas
                Excel.Range rangoMerge = excelHojaActual.Cells[RangoPrimeraCelda.Row + i, RangoPrimeraCelda.Column];
                rangoMerge = rangoMerge.Resize[1, numColumnasNuevas + 1];
                rangoMerge.Merge();

                //Recuperamos el estilo de las celdas, pues se pierde tras hacer el merge
                if (i > 2)
                {  
                    Excel.Range celdaInicialMerge = excelHojaActual.Cells[rangoMerge.Row + 1, rangoMerge.Column];
                   
                    //El + 1 es por la columna nº 1, que no cuenta. Es decir, si tengo 2 cols más, el rango va de la 1 a la 5. Column 2 + Nº Col ACtuales 2 + 1
                    Excel.Range celdaFinalMerge = excelHojaActual.Cells[rangoMerge.Row + 1, RangoPrimeraCelda.Column + numColumnasNuevas + 1];
                    
                    rangoMerge = excelHojaActual.Cells.get_Range(((Excel.Range)celdaInicialMerge.Cells[0]).Address, ((Excel.Range)celdaFinalMerge[0]).Address);                                        
                    rangoMerge.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    rangoMerge.Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 255));
                }
             
            }

        }

        /// <summary>
        /// Termina de dibujar una tabla
        /// </summary>
        public void TerminarDibujarTabla(bool EsTablaPublicorreo, Boolean EsTablaGrupoTramo = false, int? maxNumDecimales = null, Boolean EsTablaVA = false)
        {
            // Buscamos la última celda usada
            Excel.Range rango = excelHojaActual.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            // Movemos esa celda para abajo (teniendo en cuenta el margen) y la posicionamos a la izquierda (con margen de 1)
            rango = excelHojaActual.Cells[rango.Row + 1, 2];
            PrimeraCeldaTabla = rango;

            int fil = 0;
            Object[] ultimaFila = FilasTabla[FilasTabla.Count - 1];
            
            foreach (object[] colsDestino in FilasTabla)
            {
                rango = excelHojaActual.Cells[rango.Row + 1, rango.Column];
                int col = 0;
                                
                foreach (object celDestino in colsDestino)
                {
                    rango = excelHojaActual.Cells[rango.Row, 2 + ColumnaReal(col)];
                    
                    if(EsTablaGrupoTramo && !FilasTabla.Equals(ultimaFila) &&  col == 0)
                        AplicarEstilo(rango, fil, col, true, EsTablaPublicorreo);
                    //    excelHojaActual.Cells[rango.Row, 2 + ColumnaReal(0)].Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;

                    if (celDestino is string)
                    {
                        string celDestinoString = celDestino as string;
                        if (celDestinoString.Contains("%"))
                        {
                            celDestinoString = celDestinoString.Replace("%", "");
                            celDestinoString = celDestinoString.Replace(" ", "");
                            Decimal val;
                            if (Decimal.TryParse(celDestinoString, out val)) // Sin espacios y porcentajes es todavía un número
                            {
                                rango.Value2 = val / 100;
                                rango.NumberFormat = "0.00%";
                            }
                            else
                            {
                                rango.Value2 = celDestino;
                            }
                        }
                        else
                        {
                            rango.Value2 = celDestinoString;                            
                        }
                    }
                    else
                    {
                        rango.Value2 = celDestino;
                        string formatoNumero = "#,##0.0";

                        if (!maxNumDecimales.HasValue)
                        {
                            String decimales = new String('#', Redondeo - 1);
                            formatoNumero += decimales + "€";
                        }
                        else
                        {
                            formatoNumero = formatoNumero.PadRight(formatoNumero.Length + maxNumDecimales.Value - 1, '0') + "€";                          
                        }

                        rango.NumberFormat = formatoNumero;
                    }

                    if (ColumnasRequeridas[col] > 1 && col < ColumnasRequeridas.Length)
                    {
                        Excel.Range rangoMerge = excelHojaActual.Cells[rango.Row, rango.Column];
                        rangoMerge = rangoMerge.Resize[1, ColumnasRequeridas[col]];
                        rangoMerge.Merge();
                        AplicarEstilo(rangoMerge, fil, col, false, EsTablaPublicorreo);
                    }
                    else 
                    {
                        AplicarEstilo(rango, fil, col, EsTablaGrupoTramo, EsTablaPublicorreo);
                    }

                    col++;
                }
                
                fil++;
            }

            if (EsTablaVA)
            {
                DarFormatoTablaVA();
            }

            EstaDibujandoTabla = false;
        }

        public void MergeAndFit(Excel.Range r)
        {
            Excel.Range row = r.Rows[1];
            Excel.Range Column1 = r.Columns[1];
            var RangeWidth = r.Width;
            var OldColumn1Width = Column1.Width;
            
            for (int i = 1; i <= 3; i++)
            {
                Column1.ColumnWidth = RangeWidth / Column1.Width * Column1.ColumnWidth;
            }

            r.WrapText = true;
            r.MergeCells = false;

            var OldRowHeight = row.RowHeight;
            row.AutoFit();
            var FitRowHeight = row.RowHeight;
            r.MergeCells = true;
            Column1.ColumnWidth = OldColumn1Width;
            //Column1.ColumnWidth = OldColumn1Width;
            //Column1.ColumnWidth = OldColumn1Width;
            row.RowHeight = (FitRowHeight > OldRowHeight) ? FitRowHeight : OldRowHeight;         

        }
        
        /// <summary>
        /// Añade nuevas filas para la modalidad descuento por destino
        /// </summary>
        /// <param name="numDestinos">Número de destinos a pintar</param>
        /// <param name="esPublicorreo">Si la oferta es de publicorreo se pinta diferente</param>
        public void DarFormatoTablaDescuentoDestino(int numDestinos, bool esInformeTarifas, bool esPublicorreo = false)
        {
            //JCNS. LOGO. Pongo el color corporativo
            //var colorFondoAzul = System.Drawing.Color.FromArgb(0, 69, 125);
            var colorFondoAzul = System.Drawing.Color.FromArgb(0, 36, 86);
            var colorBordeAzul = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(149, 179, 215));
            var colorLetraBlanco = System.Drawing.Color.FromArgb(255, 255, 255);
            Excel.Range primeraCelda = excelHojaActual.Cells[PrimeraCeldaTabla.Row + 1,  ColumnaReal(1)];
            Excel.Range ultimaCelda = excelHojaActual.Cells[PrimeraCeldaTabla.Row + 1, ColumnaReal(numDestinos + 1) + 1];
            
            //Damos formato a la primera fila. Le ponemos azul de color de fondo y le añadimos bordes.
            Excel.Range primeraFila = excelHojaActual.get_Range(primeraCelda, ultimaCelda);            
            primeraFila.Cells.Interior.Color = colorFondoAzul;
            primeraFila.Cells.Font.Color = colorLetraBlanco;
            
            //if(!esPublicorreo)
            //    primeraCelda = excelHojaActual.Cells[PrimeraCeldaTabla.Row + 1, ColumnaReal(2) - 1];
            //else
            primeraCelda = excelHojaActual.Cells[PrimeraCeldaTabla.Row + 1, ColumnaReal(2)];

            primeraFila = excelHojaActual.get_Range(primeraCelda, ultimaCelda);
            primeraFila.Cells.Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(149, 179, 215));

            //Damos formato a la segunda fila
            primeraCelda = excelHojaActual.Cells[PrimeraCeldaTabla.Row + 2, ColumnaReal(2)];
            ultimaCelda = excelHojaActual.Cells[PrimeraCeldaTabla.Row + 2, ColumnaReal(numDestinos + 1) + 1];
            primeraFila = excelHojaActual.get_Range(primeraCelda, ultimaCelda);
            primeraFila.Cells.Interior.Color = colorFondoAzul;
            primeraCelda = excelHojaActual.Cells[PrimeraCeldaTabla.Row + 2, ColumnaReal(2) - 1];
            primeraFila = excelHojaActual.get_Range(primeraCelda, ultimaCelda);
            primeraFila.Cells.Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(149, 179, 215));
            primeraFila.Cells.Font.Color = colorLetraBlanco;

            //La primera celda de la segunda fila tiene un formato diferente
            excelHojaActual.Cells[PrimeraCeldaTabla.Row + 2, 2].MergeArea.Interior.Color = System.Drawing.Color.White;
            excelHojaActual.Cells[PrimeraCeldaTabla.Row + 2, 2].MergeArea.Borders.Color = System.Drawing.Color.White;
            
            //Nos aseguramos de que la segunda fila (Descripción de destinos) tenga la letra más pequeña
            primeraCelda = excelHojaActual.Cells[PrimeraCeldaTabla.Row + 2, ColumnaReal(2)];
            Excel.Range segundaFila = excelHojaActual.get_Range(primeraCelda, ultimaCelda);
            excelHojaActual.Rows[PrimeraCeldaTabla.Row + 2].Font.Size = 8;
                  
            String txtMasLargo = "";
            object maxWidth = 0;
            int contador = 1;

           //Obtenemos el texto mas largo
            for (int i = 0; i < segundaFila.Value2.Length; i++)
            {
                var item = segundaFila[i].Value2;
              
                if(item != null)
                {
                    String itemTxt = (String) item;
                    int longTexto = itemTxt.Length;

                    if ((longTexto > txtMasLargo.Length))
                    {
                        txtMasLargo = itemTxt;
                        maxWidth = excelHojaActual.Cells[PrimeraCeldaTabla.Row + 2, ColumnaReal(contador)].ColumnWidth * ColumnasRequeridas[contador];

                    }
                    contador++;
                }
            }

            //Asignamos a la primer celda el texto mas largo y su ancho. Al no ser una celda mergeada, vemos la altura que obtiene con autofit y esa altura
            //la usamos luego para toda la fila
            Excel.Range celdaDescripcion = excelHojaActual.Cells[PrimeraCeldaTabla.Row + 2, 1];
            Double oldColWidth = (Double)celdaDescripcion.ColumnWidth;
            Excel.Range rangoCeldaDescripcion = excelHojaActual.get_Range(celdaDescripcion, celdaDescripcion);
            rangoCeldaDescripcion.Locked = false;
            rangoCeldaDescripcion.Value2 = txtMasLargo;
            rangoCeldaDescripcion.EntireRow.WrapText = true;
            rangoCeldaDescripcion.ColumnWidth = maxWidth;
            rangoCeldaDescripcion.EntireRow.AutoFit();

            Double rowHeight = (Double)rangoCeldaDescripcion.EntireRow.RowHeight;
            rangoCeldaDescripcion.Value2 = String.Empty;

            excelHojaActual.Rows[PrimeraCeldaTabla.Row + 2].RowHeight = rowHeight;            
            segundaFila.EntireRow.WrapText = true;

            rangoCeldaDescripcion.ColumnWidth = oldColWidth;
            rangoCeldaDescripcion.Locked = true;                      
                       
            //Centramos la alineación de la tercera fila.
            if (!esInformeTarifas)
            {
                Excel.Range terceraFila = excelHojaActual.Rows[PrimeraCeldaTabla.Row + 3];
                primeraCelda = excelHojaActual.Cells[PrimeraCeldaTabla.Row + 3, 2];
                ultimaCelda = excelHojaActual.Cells[PrimeraCeldaTabla.Row + 3, ColumnaReal(numDestinos + 1) + 1];
                terceraFila = excelHojaActual.get_Range(primeraCelda, ultimaCelda);
                terceraFila.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                terceraFila.Borders.Color = colorLetraBlanco;
                terceraFila.Cells.Interior.Color = colorFondoAzul;
                terceraFila.Cells.Font.Color = colorLetraBlanco;
                //Damos formato a la tercera fila. Color de fondo, margenes, negrita, centrado y merge de las últimas celdas.

                Excel.Range cuartaFila = excelHojaActual.Rows[PrimeraCeldaTabla.Row + 4];
                primeraCelda = excelHojaActual.Cells[PrimeraCeldaTabla.Row + 4, 2];
                ultimaCelda = excelHojaActual.Cells[PrimeraCeldaTabla.Row + 4, ColumnaReal(numDestinos + 1) + 1];
                cuartaFila = excelHojaActual.get_Range(primeraCelda, ultimaCelda);
                cuartaFila.Interior.Color = colorFondoAzul;
                cuartaFila.Font.Bold = true;
                cuartaFila.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;


                primeraCelda = excelHojaActual.Cells[PrimeraCeldaTabla.Row + 4, ColumnaReal(2)];
                cuartaFila = excelHojaActual.get_Range(primeraCelda, ultimaCelda);
                cuartaFila.Merge();
                cuartaFila.Borders.Color = colorLetraBlanco;
                cuartaFila.Cells.Interior.Color = colorFondoAzul;
                cuartaFila.Cells.Font.Color = colorLetraBlanco;
            }


            //Nos aseguremos del color de la primera celda, ya que tras el formateo a veces y dependiendo de la versión de excel, se queda en azul
            primeraCelda = excelHojaActual.Cells[PrimeraCeldaTabla.Row + 1, ColumnaReal(1)];

            if (primeraCelda.Interior.Color == System.Drawing.ColorTranslator.ToOle(colorFondoAzul))
            {
                primeraCelda.Interior.Color = System.Drawing.Color.White;
            }
                
            if (primeraCelda.Value == "Peso")
            {
                primeraCelda.Font.Size = 11;
            }

            if (!esInformeTarifas)
            {
                Excel.Range cuartaFilaAux = excelHojaActual.Rows[PrimeraCeldaTabla.Row + 4];
                cuartaFilaAux.Cells.Font.Color = colorLetraBlanco;
            }
                
        }

        /// <summary>
        /// Añade nuevas filas para la modalidad descuento por destino
        /// </summary>
        /// <param name="numDestinos">Número de destinos a pintar</param>
        /// <param name="esPublicorreo">Si la oferta es de publicorreo se pinta diferente</param>
        public void DarFormatoTablaPrecioCierto(int numDestinos, int numTramos, bool esPublicorreo = false)
        {
            //JCNS. LOGO. Pongo el color corporativo
            //var colorFondoAzul = System.Drawing.Color.FromArgb(0, 69, 125);
            var colorFondoAzul = System.Drawing.Color.FromArgb(0, 36, 86);
            var colorBordeAzul = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(149, 179, 215));
            var colorLetraBlanco = System.Drawing.Color.FromArgb(255, 255, 255);
            Excel.Range primeraCelda;

            if(esPublicorreo)// || ColumnaReal(1) - ColumnaReal(0) < 3)
                primeraCelda = excelHojaActual.Cells[PrimeraCeldaTabla.Row + 1, ColumnaReal(1) + 1];
            else
                primeraCelda = excelHojaActual.Cells[PrimeraCeldaTabla.Row + 1, ColumnaReal(1)];

            Excel.Range ultimaCelda = excelHojaActual.Cells[PrimeraCeldaTabla.Row + 1, ColumnaReal(numDestinos + 1) + 1];

            //Damos formato a la primera fila. Le ponemos azul de color de fondo y le añadimos bordes.
            Excel.Range primeraFila = excelHojaActual.get_Range(primeraCelda, ultimaCelda);
            primeraFila.Cells.Interior.Color = colorFondoAzul;
            primeraFila.Cells.Font.Color = colorLetraBlanco;
            
            //if (ColumnaReal(1) - ColumnaReal(0) < 3)
            //    primeraCelda = excelHojaActual.Cells[PrimeraCeldaTabla.Row + 1, ColumnaReal(2) + 1];
            //else
                primeraCelda = excelHojaActual.Cells[PrimeraCeldaTabla.Row + 1, ColumnaReal(2)];

            
            primeraFila = excelHojaActual.get_Range(primeraCelda, ultimaCelda);
            primeraFila.Cells.Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(149, 179, 215));

            //Damos formato a la segunda fila
            primeraCelda = excelHojaActual.Cells[PrimeraCeldaTabla.Row + 2, ColumnaReal(2)];
            ultimaCelda = excelHojaActual.Cells[PrimeraCeldaTabla.Row + 2, ColumnaReal(numDestinos + 1) + 1];
            primeraFila = excelHojaActual.get_Range(primeraCelda, ultimaCelda);
            primeraFila.Cells.Interior.Color = colorFondoAzul;
            primeraCelda = excelHojaActual.Cells[PrimeraCeldaTabla.Row + 2, ColumnaReal(2) - 1];
            primeraFila = excelHojaActual.get_Range(primeraCelda, ultimaCelda);
            primeraFila.Cells.Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(149, 179, 215));
            primeraFila.Cells.Font.Color = colorLetraBlanco;

            //La primera celda de la segunda fila tiene un formato diferente
            excelHojaActual.Cells[PrimeraCeldaTabla.Row + 2, 2].MergeArea.Interior.Color = System.Drawing.Color.White;
            excelHojaActual.Cells[PrimeraCeldaTabla.Row + 2, 2].MergeArea.Borders.Color = System.Drawing.Color.White;

            //Nos aseguramos de que la segunda fila (Descripción de destinos) tenga la letra más pequeña
            primeraCelda = excelHojaActual.Cells[PrimeraCeldaTabla.Row + 2, ColumnaReal(2)];
            Excel.Range segundaFila = excelHojaActual.get_Range(primeraCelda, ultimaCelda);
            excelHojaActual.Rows[PrimeraCeldaTabla.Row + 2].Font.Size = 8;
                      
            String txtMasLargo = "";
            object maxWidth = 0;
            int contador = 1;

            //Obtenemos el texto mas largo
            for (int i = 0; i < segundaFila.Value2.Length; i++)
            {
                var item = segundaFila[i].Value2;

                if (item != null)
                {
                    String itemTxt = (String)item;
                    int longTexto = itemTxt.Length;

                    if ((longTexto > txtMasLargo.Length))
                    {
                        txtMasLargo = itemTxt;
                        maxWidth = excelHojaActual.Cells[PrimeraCeldaTabla.Row + 2, ColumnaReal(contador)].ColumnWidth * ColumnasRequeridas[contador];

                    }
                    contador++;
                }
            }

            //Asignamos a la primer celda el texto mas largo y su ancho. Al no ser una celda mergeada, vemos la altura que obtiene con autofit y esa altura
            //la usamos luego para toda la fila
            Excel.Range celdaDescripcion = excelHojaActual.Cells[PrimeraCeldaTabla.Row + 2, 1];
            Double oldColWidth = (Double)celdaDescripcion.ColumnWidth;
            Excel.Range rangoCeldaDescripcion = excelHojaActual.get_Range(celdaDescripcion, celdaDescripcion);
            rangoCeldaDescripcion.Locked = false;
            rangoCeldaDescripcion.Value2 = txtMasLargo;
            rangoCeldaDescripcion.EntireRow.WrapText = true;
            rangoCeldaDescripcion.ColumnWidth = maxWidth;
            rangoCeldaDescripcion.EntireRow.AutoFit();

            Double rowHeight = (Double)rangoCeldaDescripcion.EntireRow.RowHeight;
            rangoCeldaDescripcion.Value2 = String.Empty;

            excelHojaActual.Rows[PrimeraCeldaTabla.Row + 2].RowHeight = rowHeight;
            segundaFila.EntireRow.WrapText = true;

            rangoCeldaDescripcion.ColumnWidth = oldColWidth;
            rangoCeldaDescripcion.Locked = true;

            //Nos aseguramos de que la primera fila sea blanca            
            excelHojaActual.Cells[PrimeraCeldaTabla.Row + 1, 2].MergeArea.Interior.Color = System.Drawing.Color.White;
            excelHojaActual.Cells[PrimeraCeldaTabla.Row + 1, 2].MergeArea.Borders.Color = System.Drawing.Color.White;

            //Nos aseguremos del color de la primera celda, ya que tras el formateo a veces y dependiendo de la versión de excel, se queda en azul
            primeraCelda = excelHojaActual.Cells[PrimeraCeldaTabla.Row + 1, ColumnaReal(1)];
            if (primeraCelda.Interior.Color == System.Drawing.ColorTranslator.ToOle(colorFondoAzul))
            {
                primeraCelda.Interior.Color = System.Drawing.Color.White;
            }
                                                         
            //Quitamos negrita del literal "Descuento aplicado sobre tarifa" cuando exista
            primeraCelda = excelHojaActual.Cells[PrimeraCeldaTabla.Row + 2, 2];
            if(primeraCelda.Value == "Descuento aplicado sobre tarifa\t")
            {
                primeraCelda.Font.Bold = false;
                primeraCelda = excelHojaActual.Cells[PrimeraCeldaTabla.Row + 1, 2];
                primeraCelda.Font.Size = 11;              
            }

            if (primeraCelda.Value == "Peso")
            {
                primeraCelda.Font.Size = 11;     
            }

        }

        /// <summary>
        /// Añade nuevas filas para la modalidad descuento por destino
        /// </summary>
        /// <param name="numDestinos">Número de destinos a pintar</param>
        /// <param name="esPublicorreo">Si la oferta es de publicorreo se pinta diferente</param>
        public void DarFormatoTablaVA()
        {
            //JCNS. LOGO. Pongo el color corporativo
            //var colorFondoAzul = System.Drawing.Color.FromArgb(0, 69, 125);
            var colorFondoAzul = System.Drawing.Color.FromArgb(0, 36, 86);
            var colorBordeAzul = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(149, 179, 215));
            var colorLetraBlanco = System.Drawing.Color.FromArgb(255, 255, 255);
            Excel.Range primeraCelda;

            
            primeraCelda = excelHojaActual.Cells[PrimeraCeldaTabla.Row + 1, ColumnaReal(1) - 2];
            Excel.Range ultimaCelda = excelHojaActual.Cells[PrimeraCeldaTabla.Row + 1, ColumnaReal(1) + 2];

            //Damos formato a la primera fila. Le ponemos azul de color de fondo y le añadimos bordes.
            Excel.Range primeraFila = excelHojaActual.get_Range(primeraCelda, ultimaCelda);
            primeraFila.Cells.Interior.Color = colorFondoAzul;
            primeraFila.Cells.Font.Color = colorLetraBlanco;

        }



        /// <summary>
        /// Pone la columna indicada entera en negrita de la última tabla dibujada (Hay que llamar antes de TerminarDibujarTabla())
        /// </summary>
        /// <param name="columna"></param>
        public void PonerColumnaEnNegrita(int columna)
        {
            ColumnasNegrita.Add(columna);
        }

        /// <summary>
        /// Añade fila en la primera posición no usada
        /// </summary>
        /// <param name="colsDestino">Datos a introducir</param>
        public void AgregarFila(object[] colsDestino, bool esAnchoFijo = false, bool ajustarTextoColumna = true, bool obviarPrimeraColumna = false)
        {
            Boolean esPrimeraEjecucion = false;

            if (!EstaDibujandoTabla)
            {
                throw new Exception("IniciarDibujarTabla() antes de agregar filas");
            }

            if (NumColumnas == -1)
            {
                NumColumnas = colsDestino.Length;
                ColumnasRequeridas = new int[NumColumnas];
                esPrimeraEjecucion = true;
            }

            int columna = 0;
            //foreach (object celda in colsDestino)
            for(int i = 0; i < colsDestino.Length; i++)
            {
                Object celda = colsDestino[i];
                String enString = Convert.ToString(celda);
                double ancho = CalcularAnchoTexto(enString);
                
                if (esAnchoFijo)
                    ancho = 10;

                //[MMUNOZ] Ajusta el texto para que las columnas no sean muy anchas. 
                //Incorporamos obviarPrimeraColumna si no queremos que se añadan saltos de lineas en la primera columna
                if (esPrimeraEjecucion && ajustarTextoColumna && (!obviarPrimeraColumna || i > 0))
                {
                    if (ancho > 10)
                    {
                        int breakLineNumber = (int)ancho / 11;

                        for (int j = 1; j <= breakLineNumber; j++)
                        {
                            enString = enString.Insert(10 * j + ((j - 1) * 2), "\n");                             
                        }

                        colsDestino[i] = enString + "\n";
                        ancho = 10;
                    }
                }

                int columnasNecesarias = Convert.ToInt32(System.Math.Ceiling(ancho / AnchoCeldaBase));
                if (ColumnasRequeridas[columna] < columnasNecesarias)
                {
                    ColumnasRequeridas[columna] = columnasNecesarias;
                }
                columna++;
            }

            FilasTabla.Add(colsDestino);
        }

        /// <summary>
        /// Calcular el ancho en puntos de un texto. Asume Calibri
        /// </summary>
        /// <param name="cadena">Texto a calcular</param>
        /// <returns>Ancho</returns>
        private double CalcularAnchoTexto(String cadena)
        {
            Decimal numero;
            bool isNum = false;
            if (Decimal.TryParse(cadena, out numero)) // Es numerico
            {
                isNum = true;
                if (TamOcupado.ContainsKey(cadena.Length))
                {
                    return TamOcupado[cadena.Length];
                }
            }
            Excel.Range rango = excelHojaActual.Range["A1"];
            double anchoOriginal = rango.ColumnWidth;
            rango.Value2 = cadena;
            rango.Columns.AutoFit();
            double anchoRequerido = rango.ColumnWidth;
            rango.Columns.ColumnWidth = anchoOriginal;
            rango.Value2 = null;
            if (isNum)
            {
                TamOcupado.Add(cadena.Length, anchoRequerido);
            }
            return anchoRequerido;
        }

        #endregion

        #region Guardar el Libro

        /// <summary>
        /// Guarda el libro como
        /// </summary>
        /// <param name="nomFich"></param>
        public void GuardarComoLibro(string nomFich)
        {
            // Guardar el libro
            this.excelLibro.SaveAs(nomFich, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, this.m_objOpt,
                this.m_objOpt, this.m_objOpt, this.m_objOpt, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                this.m_objOpt, this.m_objOpt, this.m_objOpt, this.m_objOpt, this.m_objOpt);
        }

        /// <summary>
        /// Guarda el libro
        /// </summary>
        public void GuardarLibro()
        {
            this.excelHojaActual.Cells[1, 1].Select();
            

            // Guardar el libro
            this.excelLibro.Save();
        }

        #endregion

        #region Cerrar la Excel

        /// <summary>
        /// Cierra la aplicacion Excel
        /// </summary>
        public void CerrarExcel()
        {
            this.excelLibro.Close(false, this.m_objOpt, this.m_objOpt);
            this.excel.Excel.Quit();
            this.abierto = false;
        }

        #endregion

        #endregion

        #region Concatenar Excels

        /// <summary>
        /// To combine multiple workbooks into a file
        /// </summary>
        /// <remarks>
        /// The following file name convention will be used while combining the child files
        /// exportFileKey_[Description] where the description will be the tab name
        ///
        /// The ordering of the worksheets will be using the file creation time
        ///
        /// The above convention can be enhanced when necessary but need to make sure the backward compatibility for the existing codes
        ///
        /// Note:
        /// - the index starts from 1 in the excel automation array
        /// - be careful when making changes, especially moving things around in the method, e.g. prompts might come up unexpectedly
        /// - to avoid "zombie" excel instances in the task manager when referencing the COM object, please refer to the http://support.microsoft.com/default.aspx/kb/317109
        ///
        /// </remarks>
        /// <param name="exportFilePath">the destination file name choosen by the user</param>
        /// <param name="exportFileKey">the unique key file name choosen by the user, this is to avoid merging files with similar names</param>
        /// <param name="rawFilesDirectory">the folder where the files are being generated, this can be temp folder or any folder basically</param>
        /// <param name="deleteRawFiles">delete the raw files after completed?</param>
        ///
        /// <returns></returns>
        public static bool CombineWorkBooks(string exportFilePath,string[] filesToMerge, bool deleteRawFiles = false)
        {
            Application xlApp = null;
            Workbooks newBooks = null;
            Workbook newBook = null;
            Sheets newBookWorksheets = null;
            Worksheet defaultWorksheet = null;
            // IEnumerable<string> filesToMerge = null;
            //bool areRowsTruncated = false;


            //JCNS. Excel 2010 mal instalado.
            bool bExcel2010Error = false;
            try
            {
                System.Console.WriteLine("Method: CombineWorkBooks - Starting excel");
                xlApp = new Application();

                if (xlApp == null)
                {
                    System.Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
                    return false;
                }





                System.Console.WriteLine("Method: CombineWorkBooks - Disabling the display alerts to prevent any prompts during workbooks close");
                // not an elegant solution? however has to do this else will prompt for save on exit, even set the Saved property didn't help
                xlApp.DisplayAlerts = false;

                // Console.WriteLine("Method: CombineWorkBooks - Set Visible to false as a background process, else it will be displayed in the task bar");
                xlApp.Visible = false;

#if DEBUG
                if (ConfigurationManager.AppSettings["OfficeAppVisible"] == "S")
                {
                    xlApp.Visible = true;
                }
#endif                



                //Console.WriteLine("Method: CombineWorkBooks - Create a new workbook, comes with an empty default worksheet");
                newBooks = xlApp.Workbooks;
                newBook = newBooks.Add(XlWBATemplate.xlWBATWorksheet);
                newBookWorksheets = newBook.Worksheets;

                xlApp.Visible = true;

                // get the reference for the empty default worksheet
                if (newBookWorksheets.Count > 0)
                {
                    defaultWorksheet = newBookWorksheets[1] as Worksheet;
                }

                // Console.WriteLine("Method: CombineWorkBooks - Get the files sorted by creation date");
                //var dirInfo = new DirectoryInfo(rawFilesDirectory);
                //  filesToMerge = from f in dirInfo.GetFiles(exportFileKey + "*", SearchOption.TopDirectoryOnly)
                //                orderby f.CreationTimeUtc
                //               select f.FullName;

                foreach (var filePath in filesToMerge)
                {
                    Workbook childBook = null;
                    Sheets childSheets = null;
                    try
                    {
                        System.Console.WriteLine("Method: CombineWorkBooks - Processing {0}", filePath);
                        childBook = newBooks.Open(filePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing
                                                             , Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                        childSheets = childBook.Worksheets;
                        if (childSheets != null)
                        {
                            for (int iChildSheet = 1; iChildSheet <= childSheets.Count; iChildSheet++)
                            {
                                Worksheet sheetToCopy = null;
                                try
                                {
                                    sheetToCopy = childSheets[iChildSheet] as Worksheet;
                                    if (sheetToCopy != null)
                                    {
                                        System.Console.WriteLine("Method: CombineWorkBooks - Assigning the worksheet name");
                                        //sheetToCopy.Name = Truncate(GetReportDescription(Path.GetFileNameWithoutExtension(filePath), sheetToCopy.Name), 31); // only 31 char max

                                        System.Console.WriteLine("Method: CombineWorkBooks - Copy the worksheet before the default sheet");
                                        sheetToCopy.Copy(defaultWorksheet, Type.Missing);
                                    }
                                }
                                catch (Exception ex1)
                                {
                                    string error = ex1.Message;
                                }
                                finally
                                {
                                    DisposeCOMObject(sheetToCopy);
                                }
                            }

                            System.Console.WriteLine("Method: CombineWorkBooks - Close the childbook");
                            // for some reason, calling close below may cause an exception -> System.Runtime.InteropServices.COMException (0x80010108): The object invoked has disconnected from its clients.
                            childBook.Close(false, Type.Missing, Type.Missing);
                        }
                    }
                    finally
                    {
                        DisposeCOMObject(childSheets);
                        DisposeCOMObject(childBook);
                    }
                }

                // Console.WriteLine("Method: CombineWorkBooks - Delete the empty default worksheet");
                if (defaultWorksheet != null) defaultWorksheet.Delete();




                //Excel.XlFileFormat.xlWorkbookDefault

                // Console.WriteLine("Method: CombineWorkBooks - Save the new book into the export file path: {0}", exportFilePath);
                //...........................................................................................................................
                //JCNS. ¡¡ O J O !! ME ESTÁ DANDO ERROR AL GRABAR, PERO EN OTROS EQUIPOS FUNCIONA. ALGO ESTÁ MAL INSTALADO DEL EXCEL 2010
                //...........................................................................................................................

                newBook.SaveAs(exportFilePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing
                    , XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                //newBook.SaveAs(exportFilePath, XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, Type.Missing, Type.Missing
                //        , XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                //newBook.SaveAs(exportFilePath, Excel.XlFileFormat.xlOpenXMLWorkbook, Missing.Value, Missing.Value, false, false
                //    , XlSaveAsAccessMode.xlNoChange
                //    , XlSaveConflictResolution.xlUserResolution
                //    , true, Missing.Value, Missing.Value, Missing.Value);

                newBooks.Close();
                xlApp.DisplayAlerts = true;

                return true;
            }
            catch (Exception ex)
            {
                //JCNS. Excel 2010 mal instalado.
                bExcel2010Error = true;
                
                System.Console.WriteLine("Method: CombineWorkBooks - Exception: {0}", ex.ToString());
                return false;
            }
            finally
            {
                DisposeCOMObject(defaultWorksheet);
                DisposeCOMObject(newBookWorksheets);
                DisposeCOMObject(newBooks);
                DisposeCOMObject(newBook);

                //  Console.WriteLine("Method: CombineWorkBooks - Closing the excel app");
                if (xlApp != null)
                {
                    if (bExcel2010Error == false)
                    {
                        xlApp.Quit();
                    }
                    DisposeCOMObject(xlApp);
                }

                if (deleteRawFiles)
                {
                    //JCNS. Excel 2010 mal instalado.
                    
                    if (bExcel2010Error == false)
                    {
                        System.Console.WriteLine("Method: CombineWorkBooks - Deleting the temporary files");
                        DeleteTemporaryFiles(filesToMerge);
                    }
                }
            }
        }

        public static bool CombineWorkBooks_version14(string exportFilePath, string[] filesToMerge, bool deleteRawFiles = false)
        {
            Application xlApp = null;
            Workbooks newBooks = null;
            Workbook newBook = null;
            Sheets newBookWorksheets = null;
            Worksheet defaultWorksheet = null;
            // IEnumerable<string> filesToMerge = null;
            //bool areRowsTruncated = false;

            try
            {
                ManagerExcel objExcel = new ManagerExcel(exportFilePath, true);

                if (!File.Exists(exportFilePath))
                {

                    //Copiamos la plantilla en el fichero temporal.
                    File.Copy(filesToMerge[0], exportFilePath, true);
                }
                try
                {
                    objExcel.AbrirFichero();
                }
                catch { }
                if (objExcel.Abierto)
                {
                    
                    //objExcel.SeleccionarPrimeraHojaLibreOCrearNuevaYRenombrar(objProductoBE.CodAnexoSAP + " - " + objProductoBE.CodProducto + " - " + GenerarAbreviaturaModeloNegociacion(objProductoOfertaBE.CodModalidadNegociacion.Trim()));
                }

                if (objExcel.Abierto)
                {
                    objExcel.GuardarLibro();
                    objExcel.CerrarExcel();
                }

                
                    ManagerExcel.AbrirExcelStandalone(exportFilePath);
   















                return true;
            }
            catch (Exception ex)
            {
                //JCNS. LOGO
                xlApp.Visible = true;
                System.Console.WriteLine("Method: CombineWorkBooks - Exception: {0}", ex.ToString());
                return false;
            }
            finally
            {
                DisposeCOMObject(defaultWorksheet);
                DisposeCOMObject(newBookWorksheets);
                DisposeCOMObject(newBooks);
                DisposeCOMObject(newBook);

                //  Console.WriteLine("Method: CombineWorkBooks - Closing the excel app");
                if (xlApp != null)
                {
                    xlApp.Quit();
                    DisposeCOMObject(xlApp);
                }

                if (deleteRawFiles)
                {
                    System.Console.WriteLine("Method: CombineWorkBooks - Deleting the temporary files");
                    DeleteTemporaryFiles(filesToMerge);
                }
            }
        }

        private static void DisposeCOMObject(object o)
        {
            System.Console.WriteLine("Method: DisposeCOMObject - Disposing");
            if (o == null)
            {
                return;
            }
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(o);
            }
            catch (Exception ex)
            {
                System.Console.WriteLine("Method: DisposeCOMObject - Exception: {0}", ex.ToString());
            }
        }

        private static void DeleteTemporaryFiles(IEnumerable<string> tempFilenames)
        {
            foreach (var tempFile in tempFilenames)
            {
                try
                {
                    File.Delete(tempFile);
                }
                catch
                    (Exception)
                {
                    System.Console.WriteLine("Could not delete temporary file '{0}'", tempFilenames);
                }
            }
        }

        /// <summary>
        /// the first array item will be the key
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="defaultName"></param>
        /// <returns></returns>
        protected static string GetReportDescription(string fileName, string defaultName)
        {
            var splits = fileName.Split('_');
            return splits.Length > 1 ? string.Join("-", splits, 1, splits.Length - 1) : defaultName;
        }

        /// <summary>
        /// Get a substring of the first N characters.
        /// http://dotnetperls.com/truncate-string
        /// </summary>
        public static string Truncate(string source, int length)
        {
            if (source.Length > length)
            {
                source = source.Substring(0, length);
            }
            return source;
        }

        #endregion 
        
    }
}

