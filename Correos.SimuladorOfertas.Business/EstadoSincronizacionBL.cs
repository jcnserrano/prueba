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
    public class EstadoSincronizacionBL
    {
        #region Metodos Insert

        /// <summary>
        /// Inserta un nuevo registro en base de datos
        /// </summary>
        /// <param name="idEstado">ID del estado, recibido desde SAP CRM</param>
        /// <param name="idOferta">ID de la oferta a consultar</param>
        /// <param name="estado"></param>
        public void InsertEstadoSincronizacion(Guid idEstado, Guid idOferta, String estado, DateTime fechaInicio)
        {
            var item = new EstadoSincronizacionBE(idEstado, idOferta, estado, fechaInicio);
            InsertEstadoSincronizacion(item);
        }

        /// <summary>
        /// Inserta un nuevo registro en base de datos
        /// </summary>
        /// <param name="estadoSincronizacion">Objeto a incluir en BBDD</param>        
        public void InsertEstadoSincronizacion(EstadoSincronizacionBE estadoSincronizacion)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                var estadoPersistence = new EstadoSincronizacionPersistence(uow);
                estadoPersistence.InsertEstadoSincronizacion(estadoSincronizacion);
                uow.Save();
            }
        }
        
        #endregion 

        #region Métodos Obtener

        /// <summary>
        /// Obtiene todos los estados de sincronización de las ofertas
        /// </summary>
        /// <returns></returns>
        public List<EstadoSincronizacionBE> ObtenerEstadosSincronizacion()
        {
            List<EstadoSincronizacionBE> resultado = new List<EstadoSincronizacionBE>();

            using (IUnitOfWork uow = new UnitOfWork())
            {
                EstadoSincronizacionPersistence estadoPersistence = new EstadoSincronizacionPersistence(uow);
                resultado = estadoPersistence.ObtenerEstadosSincronizacion();
                uow.Save();
            }

            return resultado;
        }

        /// <summary>
        /// Obtiener el estado de sincronización de una oferta
        /// </summary>
        /// <param name="idOferta"></param>
        /// <returns></returns>
        public EstadoSincronizacionBE ObtenerEstadoSincronizacionOferta(Guid idOferta)
        {
            EstadoSincronizacionBE resultado = null;

            using(IUnitOfWork uow = new UnitOfWork())
            {
                EstadoSincronizacionPersistence estadoPersistence = new EstadoSincronizacionPersistence(uow);
                resultado = estadoPersistence.ObtenerEstadoSincronizacionOferta(idOferta);
                uow.Save();
            }

            return resultado;
        }

        /// <summary>
        /// Obtiene si hay ofertas sincronizándose
        /// </summary>
        /// <param name="idOferta"></param>
        /// <returns></returns>
        public Boolean HayOfertasSincronizando()
        {
            Boolean resultado = false;

            using (IUnitOfWork uow = new UnitOfWork())
            {
                EstadoSincronizacionPersistence estadoPersistence = new EstadoSincronizacionPersistence(uow);
                resultado = estadoPersistence.HayOfertasSincronizando();
                uow.Save();
            }

            return resultado;
        }

        #endregion

        #region Métodos Actualizar 

        /// <summary>
        /// Modifica el estado de sincronización de una oferta
        /// </summary>
        /// <param name="idOferta"></param>
        /// <param name="idEstado"></param>
        /// <param name="estado">Estado de Sinc. de la oferta: En Desarrollo, Error, Finalizada</param>
        public void ActualizarEstadoSincronizacion(Guid idOferta, Guid idEstado, String estado)
        {
            using(IUnitOfWork uow = new UnitOfWork())
            {
                EstadoSincronizacionPersistence estadoPersistence = new EstadoSincronizacionPersistence(uow);
                estadoPersistence.ActualizarEstadoSincronizacion(idOferta, idEstado, estado);
                uow.Save();                
            }          
        }

        #endregion

        #region Metodos Eliminar

        /// <summary>
        /// Elimina el estado de sincronización de una oferta
        /// </summary>
        /// <param name="idOferta">ID de la oferta de la que se quiere borrar su estado de sincronización.</param>
        public void EliminarEstadoSincronizacion(Guid idOferta)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                EstadoSincronizacionPersistence estadoPersistence = new EstadoSincronizacionPersistence(uow);
                estadoPersistence.EliminarEstadoSincronizacion(idOferta);
                uow.Save();
            }
        }    
      
        #endregion 
    }
}
