using Correos.SimuladorOfertas.DTOs;
using Correos.SimuladorOfertas.Persistence;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Correos.SimuladorOfertas.Business
{
    public class AgrupacionBL
    {
        #region OBTENER

        /// <summary>
        /// Obtiene las agrupaciones de un producto
        /// </summary>
        /// <param name="idProducto"></param>
        /// <returns></returns>
        public List<AgrupacionBE> ObtenerAgrupacionesProducto(Guid idProducto)
        {
            List<AgrupacionBE> result;

            using (IUnitOfWork uow = new UnitOfWork())
            {
                AgrupacionPersistence objPersistencia = new AgrupacionPersistence(uow);
                result = objPersistencia.ObtenerAgrupacionesProducto(idProducto);                
            }
            
            return result;
        }

        /// <summary>
        /// Obtiene los destinos asociados a una agrupación
        /// </summary>
        /// <param name="idAgrupacion"></param>
        /// <returns></returns>
        public List<AgrupacionDestinoBE> ObtenerAgrupacionesDestino(Guid idAgrupacion)
        {
            List<AgrupacionDestinoBE> result;

            using (IUnitOfWork uow = new UnitOfWork())
            {
                AgrupacionPersistence objPersistencia = new AgrupacionPersistence(uow);
                result = objPersistencia.ObtenerAgrupacionesDestino(idAgrupacion);
            }

            return result;
        }

        #endregion

        #region INSERTAR

        //public bool GuardarAgrupaciones(Collection<AgrupacionBE> listaAgrupaciones)
        //{
        //    bool result;

        //    using (IUnitOfWork uow = new UnitOfWork())
        //    {
        //        AgrupacionPersistence objPersistencia = new AgrupacionPersistence(uow);
        //        result = objPersistencia.GuardarAgrupaciones(listaAgrupaciones);
        //    }

        //    return result;
        //}

        public bool GuardarAgrupacionesDestino(List<AgrupacionDestinoBE> listaAgrupacionesDestino)
        {
            bool result;

            using (IUnitOfWork uow = new UnitOfWork())
            {
                AgrupacionPersistence objPersistencia = new AgrupacionPersistence(uow);
                result = objPersistencia.GuardarAgrupacionesDestino(listaAgrupacionesDestino);
            }

            return result;
        }

        #endregion

        #region BORRAR

        /// <summary>
        /// Método que borrar una agrupación y todos sus destinos asociados
        /// </summary>
        /// <param name="idAgrupacion"></param>
        /// <returns></returns>
        public bool BorrarAgrupacion(Guid idAgrupacion)
        {
            bool result;

            using (IUnitOfWork uow = new UnitOfWork())
            {
                AgrupacionPersistence objPersistencia = new AgrupacionPersistence(uow);
                result = objPersistencia.BorrarAgrupacion(idAgrupacion);
            }

            return result;
        }


        /// <summary>
        /// Método que borra todas las agrupaciones de un producto
        /// </summary>
        /// <param name="idAgrupacion"></param>
        /// <returns></returns>
        public bool BorrarAgrupacionesProducto(Guid idProducto)
        {
            bool result;

            using (IUnitOfWork uow = new UnitOfWork())
            {
                AgrupacionPersistence objPersistencia = new AgrupacionPersistence(uow);
                result = objPersistencia.BorrarAgrupacionesProducto(idProducto);
            }

            return result;
        }
        #endregion           

    }
}
