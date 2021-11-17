using System;
using System.Collections.ObjectModel;
using Correos.SimuladorOfertas.Common;
using Correos.SimuladorOfertas.DTOs;
using Correos.SimuladorOfertas.Persistence;
using System.Diagnostics;
using System.Linq;
using System.Collections.Generic;

namespace Correos.SimuladorOfertas.Business
{
    public class RelacionProductosBL
    {
        #region Metodos Insert

        /// <summary>
        /// Inserta un nuevo registro en la base de datos
        /// </summary>
        /// <param name="relacionProductos"></param>
        public void InsertRelacionProductos(RelacionProductosBE relacionProductos)
        {
            //alta
            using (IUnitOfWork uow = new UnitOfWork())
            {
                RelacionProductosPersistence relacionPersistence = new RelacionProductosPersistence(uow);
                relacionPersistence.InsertRelacionProductos(relacionProductos);
            }            
        }

        #endregion 

        #region Métodos Obtener

        /// <summary>
        /// Obtiene un listado de cod productos al CodProductoSAP pasado
        /// </summary>
        /// <param name="CodProductoSAP">Código del producto del que queremos buscar sus productos relacionados</param>
        /// <param name="EsCampoA">Si True, queremos buscar los asociados de un producto del Campo A, si es false, queremos buscar los productos
        /// asociados a un producto del campo B</param>
        /// <returns></returns>
        public List<String> ObtenerRelacionProductos(String CodProductoSAP, Boolean EsCampoA = true)
        {
            List<String> resultado = new List<string>();
            
            using (IUnitOfWork uow = new UnitOfWork())
            {
                RelacionProductosPersistence relacionPersistence = new RelacionProductosPersistence(uow);
                resultado = relacionPersistence.ObtenerRelacionProductos(CodProductoSAP, EsCampoA);
            }

            return resultado;
        }

        /// <summary>
        /// Obtiene un listado con los productos de devolución
        /// </summary>
        /// <returns></returns>
        public List<String> ObtenerProductosDevolucion()
        {
            List<String> resultado = new List<string>();

            using (IUnitOfWork uow = new UnitOfWork())
            {
                RelacionProductosPersistence relacionPersistence = new RelacionProductosPersistence(uow);
                resultado = relacionPersistence.ObtenerProductosDevolucion();
            }

            return resultado;
        }

        #endregion

        #region Metodos Eliminar

        /// <summary>
        /// Elimina todos los registros en la BBDD
        /// </summary>
        public void DeleteRelacionProductos()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                RelacionProductosPersistence relacionPersistence = new RelacionProductosPersistence(uow);
                relacionPersistence.DeleteRelacionProductos();
            }            
        }

        /// <summary>
        /// Borra todos los registros asociados a un producto SAP
        /// </summary>
        /// <param name="CodProductoSAP"></param>
        public void DeleteRelacionesProdSAPSobrantes(String CodProductoSAP, Collection<RelacionProductosBE> relacionesAGuardar)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                RelacionProductosPersistence relacionPersistence = new RelacionProductosPersistence(uow);
                relacionPersistence.DeleteRelacionesProdSAPSobrantes(CodProductoSAP, relacionesAGuardar);
            }
        }

        #endregion 
    }
}
