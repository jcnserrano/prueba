namespace Correos.SimuladorOfertas.Common
{
    public class AplicacionExcel
    {
        #region Miembros de la clase

        // Aplicación excel
        protected Microsoft.Office.Interop.Excel.Application excel;

        #endregion

        #region Propiedades de la clase

        // Aplicación excel
        public Microsoft.Office.Interop.Excel.Application Excel
        {
            get { return this.excel; }
        }
        #endregion

        #region Constructores
        // Implementación del patrón Singleton
        // Constructor privado para no permitir creación de múltiples instancias de la clase
        public AplicacionExcel()
        {
            // Aplicación excel
            this.excel = new Microsoft.Office.Interop.Excel.Application();
        }

        public AplicacionExcel(Microsoft.Office.Interop.Excel.Application excel)
        {
            // Aplicación excel
            this.excel = excel;
        }

        #endregion
    }
}
