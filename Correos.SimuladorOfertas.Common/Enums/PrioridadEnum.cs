namespace Correos.SimuladorOfertas.ExcelWorkbook.Enums
{
    /// <summary>
    /// Enumerado para los controles de usuario
    /// </summary>
    public enum Prioridad
    {
        Alta,
        Media,
        Baja
    }

    public static class PrioridadEnum
    {
        /// <summary>
        /// Devuelve el texto asociado a un enumerado
        /// </summary>
        /// <param name="objEnumerado"></param>
        /// <returns></returns>
        public static string ObtenerNombreEnum(Prioridad objEnumerado)
        {
            switch (objEnumerado)
            {
                case Prioridad.Alta:
                    return "Alta";

                case Prioridad.Media:
                    return "Media";

                case Prioridad.Baja:
                    return "Baja";

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
            if (objEnumerado == "Alta")
            {
                return "1";
            }
            else if (objEnumerado == "Media")
            {
                return "2";
            }
            else if (objEnumerado == "Baja")
            {
                return "3";
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
            if (objEnumerado == "1")
            {
                return ObtenerNombreEnum(Prioridad.Alta);
            }
            else if (objEnumerado == "2")
            {
                return ObtenerNombreEnum(Prioridad.Media);
            }
            else
            {
                return ObtenerNombreEnum(Prioridad.Baja);
            }
        }

    }
}
