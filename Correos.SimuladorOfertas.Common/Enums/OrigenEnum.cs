namespace Correos.SimuladorOfertas.ExcelWorkbook.Enums
{
    /// <summary>
    /// Enumerado para los controles de usuario
    /// </summary>
    public enum Origen
    {
        IniciativaGestor,
        IniciativaCliente,
        CampanaCorreos,
        EmpresaGrupo,
        CorreosOnLine,
        RedOficinas,
        CampanaProspeccionPymes
    }

    public static class OrigenEnum
    {
        /// <summary>
        /// Devuelve el texto asociado a un enumerado
        /// </summary>
        /// <param name="objEnumerado"></param>
        /// <returns></returns>
        public static string ObtenerNombreEnum(Origen objEnumerado)
        {
            switch (objEnumerado)
            {
                case Origen.IniciativaGestor:
                    return "Iniciativa Gestor";

                case Origen.IniciativaCliente:
                    return "Iniciativa del Cliente";

                case Origen.CampanaCorreos:
                    return "Campaña Correos";

                case Origen.EmpresaGrupo:
                    return "Empresa del Grupo";

                case Origen.CorreosOnLine:
                    return "Correos On Line";

                case Origen.RedOficinas:
                    return "Red de Oficinas";

                case Origen.CampanaProspeccionPymes:
                    return "Campaña Prospección Pymes";

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
            if (objEnumerado == "Iniciativa Gestor")
            {
                return "001";
            }
            else if (objEnumerado == "Iniciativa del Cliente")
            {
                return "002";
            }
            else if (objEnumerado == "Campaña Correos")
            {
                return "003";
            }
            else if (objEnumerado == "Empresa del Grupo")
            {
                return "004";
            }
            else if (objEnumerado == "Correos On Line")
            {
                return "005";
            }
            else if (objEnumerado == "Red de Oficinas")
            {
                return "006";
            }
            else if (objEnumerado == "Campaña Prospección Pymes")
            {
                return "007";
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
            if (objEnumerado == "001")
            {
                return ObtenerNombreEnum(Origen.IniciativaGestor);
            }
            else if (objEnumerado == "002")
            {
                return ObtenerNombreEnum(Origen.IniciativaCliente);
            }
            else if (objEnumerado == "003")
            {
                return ObtenerNombreEnum(Origen.CampanaCorreos);
            }
            else if (objEnumerado == "004")
            {
                return ObtenerNombreEnum(Origen.EmpresaGrupo);
            }
            else if (objEnumerado == "005")
            {
                return ObtenerNombreEnum(Origen.CorreosOnLine);
            }
            else if (objEnumerado == "006")
            {
                return ObtenerNombreEnum(Origen.RedOficinas);
            }
            else
            {
                return ObtenerNombreEnum(Origen.CampanaProspeccionPymes);
            }
        }
    }
}
