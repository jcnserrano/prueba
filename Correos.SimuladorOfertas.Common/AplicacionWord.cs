
namespace Correos.SimuladorOfertas.Common
{
    public class AplicacionWord
    {
        #region Propiedades

        /// <summary>
        /// Contiene la instancia de word
        /// </summary>
        protected Microsoft.Office.Interop.Word.Application word;
        
        /// <summary>
        /// Contiene la instancia publica de word
        /// </summary>
        public Microsoft.Office.Interop.Word.Application Word
        {
            get { return this.word; }
        }

        #endregion
        
        #region Constructores

        // Implementación del patrón Singleton
        // Constructor privado para no permitir creación de múltiples instancias de la clase
        public AplicacionWord()
        {
            // Aplicación excel
            this.word = new Microsoft.Office.Interop.Word.Application();
        }
        #endregion
    }
}
