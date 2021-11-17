using Correos.SimuladorOfertas.DTOs;
using Correos.SimuladorOfertas.Persistence;
using System.Collections.ObjectModel;

namespace Correos.SimuladorOfertas.Business
{
    public class InformacionDestinosBL
    {
        #region Metodos publicos

        /// <summary>
        /// Obtiene la lista de destinos de un productoSAP
        /// </summary>
        /// <param name="codClienteSAP"></param>
        /// <returns></returns>
        public Collection<InformacionDestinosBE> ObtenerListadoInformacionDestinos(string codProductoSAP)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                InformacionDestinosPersistence persistencia = new InformacionDestinosPersistence(uow);
                return persistencia.ObtenerListaDestinosProducto(codProductoSAP);
            }
        }

        #endregion
    }
}
