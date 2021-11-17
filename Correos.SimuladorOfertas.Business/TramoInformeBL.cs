using Correos.SimuladorOfertas.DTOs;
using Correos.SimuladorOfertas.Persistence;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;

namespace Correos.SimuladorOfertas.Business
{
    public class TramoInformeBL
    {

        /// <summary>
        /// Obtiene, para un producto, los tramos que se han de mostrar al generar un informe desde el Excel
        /// </summary>
        /// <param name="CodProductoSAP"></param>
        /// <returns></returns>
        public List<TramoInformeBE> ObtenerTramosInformeProducto(String CodProductoSAP)
        {
            List<TramoInformeBE> resultado = new List<TramoInformeBE>();

            using (IUnitOfWork uow = new UnitOfWork())
            {
                resultado = new TramoInformePersistence(uow).ObtenerTramosInformeProducto(CodProductoSAP);
            }

            return resultado;
        }      
    }
}
