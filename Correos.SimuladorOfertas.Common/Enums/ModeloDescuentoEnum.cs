namespace Correos.SimuladorOfertas.Common.Enums
{
    /// <summary>
    /// Enumerado para los controles de usuario
    /// </summary>
    public enum ModeloDescuento
    {
        Paqueteria,
        Costes,
        Tramos,
        Volumetrico,
        Publicorreo,
        Libros,
        Publicaciones,
        Etiquetas
    }

    public static class ModeloDescuentoEnum
    {
        /// <summary>
        /// Devuelve el texto asociado a un enumerado
        /// </summary>
        /// <param name="objEnumerado"></param>
        /// <returns></returns>
        public static string ObtenerNombreEnum(ModeloDescuento objEnumerado)
        {
            switch (objEnumerado)
            {
                case ModeloDescuento.Paqueteria:
                    return "Paqueteria";

                case ModeloDescuento.Costes:
                    return "Costes";

                case ModeloDescuento.Tramos:
                    return "Tramos";

                case ModeloDescuento.Volumetrico:
                    return "Volumetrico";

                case ModeloDescuento.Publicorreo:
                    return "Publicorreo";

                case ModeloDescuento.Libros:
                    return "Libros";

                case ModeloDescuento.Publicaciones:
                    return "Publicaciones";

                case ModeloDescuento.Etiquetas:
                    return "Etiquetas";

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
            if (objEnumerado == "Paqueteria")
            {
                return "1";
            }
            else if (objEnumerado == "Costes")
            {
                return "2";
            }
            else if (objEnumerado == "Tramos")
            {
                return "3";
            }
            else if (objEnumerado == "Volumetrico")
            {
                return "4";
            }
            else if (objEnumerado == "Publicorreo")
            {
                return "5";
            }
            else if (objEnumerado == "Libros")
            {
                return "6";
            }
            else if (objEnumerado == "Publicaciones")
            {
                return "7";
            }
            else if (objEnumerado == "Etiquetas")
            {
                return "8";
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
            if (objEnumerado == "Paqueteria")
            {
                return ObtenerNombreEnum(ModeloDescuento.Paqueteria);
            }
            else if (objEnumerado == "Costes")
            {
                return ObtenerNombreEnum(ModeloDescuento.Costes);
            }
            else if (objEnumerado == "Tramos")
            {
                return ObtenerNombreEnum(ModeloDescuento.Tramos);
            }
            else if (objEnumerado == "Volumetrico")
            {
                return ObtenerNombreEnum(ModeloDescuento.Volumetrico);
            }
            else if (objEnumerado == "Publicorreo")
            {
                return ObtenerNombreEnum(ModeloDescuento.Publicorreo);
            }
            else if (objEnumerado == "Libros")
            {
                return ObtenerNombreEnum(ModeloDescuento.Libros);
            }
            else if (objEnumerado == "Publicaciones")
            {
                return ObtenerNombreEnum(ModeloDescuento.Publicaciones);
            }
            else if (objEnumerado == "Etiquetas")
            {
                return ObtenerNombreEnum(ModeloDescuento.Etiquetas);
            }
            else
            {
                return string.Empty;
            }
        }

        public static System.Collections.Generic.List<System.String> GetPaqueteriasQueSonPublicorreos()
        {
            return new System.Collections.Generic.List<System.String>(){"S0150", "S0151"};
        }
    }
}
