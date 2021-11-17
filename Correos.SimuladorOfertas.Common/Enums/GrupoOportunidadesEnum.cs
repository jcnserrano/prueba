
namespace Correos.SimuladorOfertas.ExcelWorkbook.Enums
{
    /// <summary>
    /// Enumerado para los controles de usuario
    /// </summary>
    public enum GrupoOportunidades
    {
        Captacion,
        Reajuste,
        Ampliacion        
    }

    public static class GrupoOportunidadesEnum 
    {
        /// <summary>
        /// Devuelve el texto asociado a un enumerado
        /// </summary>
        /// <param name="objEnumerado"></param>
        /// <returns></returns>
        public static string ObtenerNombreEnum(GrupoOportunidades objEnumerado)
        {
            switch (objEnumerado)
            {
                case GrupoOportunidades.Ampliacion:
                    return "Ampliación de contrato";
                    
                case GrupoOportunidades.Captacion:
                    return "Captación";
                    
                case GrupoOportunidades.Reajuste:
                    return "Reajuste";
                    
                default:
                    return "";
            }
        }

        /// <summary>
        /// Devuelve el identificador asociado al enumerado
        /// </summary>
        /// <param name="objEnumerado"></param>
        /// <returns></returns>
        public static string ObtenerCodigoEnum(string objEnumerado)
        {
            if (objEnumerado == "Ampliación de contrato")
            {
                return "Z004";
            }
            else if(objEnumerado == "Captación")
            {
                return "Z001";
            }
            else if(objEnumerado == "Reajuste")
            {
                return "Z002";
            } 
            else
            {
                return "0";
            }                
        }

        /// <summary>
        /// Devuelve el Enum asociado al codigo
        /// </summary>
        /// <param name="objEnumerado"></param>
        /// <returns></returns>
        public static string ObtenerEnumCodigo(string objEnumerado)
        {
            if (objEnumerado == "Z004")
            {
                return ObtenerNombreEnum(GrupoOportunidades.Ampliacion);
            }
            else if (objEnumerado == "Z002")
            {
                return ObtenerNombreEnum(GrupoOportunidades.Reajuste);
            }
            else 
            {
                return ObtenerNombreEnum(GrupoOportunidades.Captacion);
            }            
        }
    }
}
