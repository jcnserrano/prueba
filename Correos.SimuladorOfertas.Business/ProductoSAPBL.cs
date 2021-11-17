using Correos.SimuladorOfertas.Persistence;
using System.Collections.ObjectModel;

namespace Correos.SimuladorOfertas.Business
{
    public class ProductoSAPBL
    {
        #region Metodos Publicos

        /// <summary>
        /// Devuelve todas las descripciones de los productos por cada anexo para el filtro inteligente del arbol de productos
        /// Ejemplo: "CARTA(I)-S0028-ANEXO1"
        /// </summary>
        /// <returns></returns>
        public Collection<string> ObtenerDescripcionesProductosFiltro()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProductoSAPPersistence objProductoSAPP = new ProductoSAPPersistence(uow);
                return objProductoSAPP.ObtenerDescripcionesProductosFiltro();
            }
        }

        /// <summary>
        /// Obtiene el peso volumétrico máximo de un producto o 0 si tiene que coger el valor por defecto.
        /// </summary>
        /// <returns></returns>
        public int ObtenerPesoVolumetricoMaxProducto(string codProductoSAP)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProductoSAPPersistence objProductoSAP = new ProductoSAPPersistence(uow);
                return objProductoSAP.ObtenerPesoVolumetricoMaxProducto(codProductoSAP);
            }
        }

        #endregion


    }
}
