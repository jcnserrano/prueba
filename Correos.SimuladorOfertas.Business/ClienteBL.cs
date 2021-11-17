using Correos.SimuladorOfertas.DTOs;
using Correos.SimuladorOfertas.InOutLight;
using Correos.SimuladorOfertas.Persistence;
using System;
using System.Collections.ObjectModel;

namespace Correos.SimuladorOfertas.Business
{
    public class ClienteBL
    {
        #region Métodos Publicos

        #region Obtener clientes de SAP

        /// <summary>
        /// Devuelve el listado de clientes de un determinado Gestor.
        /// </summary>
        /// <param name="usuario">Usuario del que se van a obtener sus clientes</param>
        /// <param name="numHits">Número de resultados a mostrar</param>
        /// <param name="esSeleccionable">Flag que indica si sobre los clientes mostrados se va a poder crear una nueva oferta o no</param>
        /// <param name="filtros">Filtros a aplicar en la búsqueda</param>
        /// <returns>Listado de Cliente_Ent</returns>
        public Collection<ClienteBE> ObtenerClientesDelGestor(string usuario, string password, string numHits, string esSeleccionable, string tipoCliente, Collection<ParametrosBusquedaClienteBE> filtros)
        {
            if (SSOHelper.Instance.LogarConSSO)
            {
                CommunicatorLight conectorSAP = new CommunicatorLight(SSOHelper.Instance.Usuario, password);
                //SSOHelper.Instance.ActualizarCookiePortal();
                SSOHelper.Instance.InicializarWSLight(conectorSAP.Cliente);
                Collection<ClienteBE> result = conectorSAP.ZCClientesGestorRfc(numHits, esSeleccionable, tipoCliente, filtros, usuario);
                SSOHelper.Instance.LimpiarWSLight();
                return result;
            }
            else
            {
                CommunicatorLight conectorSAP = new CommunicatorLight(usuario, password);
                return conectorSAP.ZCClientesGestorRfc(numHits, esSeleccionable, tipoCliente, filtros, usuario);
            }
        }

        #endregion

        #region Obtener clientes de BBDD

        /// <summary>
        /// Obtiene un objeto ClienteBE a partir del idCliente.
        /// </summary>
        /// <param name="codClienteSAP"></param>
        /// <returns></returns>
        public ClienteBE ObtenerClienteByIdCliente(Guid idCliente)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ClientePersistence persistencia = new ClientePersistence(uow);
                return persistencia.ObtenerClienteByIdCliente(idCliente);
            }
        }

        /// <summary>
        /// Método que obtiene todos los clientes de BD
        /// </summary>
        /// <returns>Colección de entidades ClienteBE</returns>
        public Collection<ClienteBE> ObtenerClientesBD()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ClientePersistence persistencia = new ClientePersistence(uow);
                return persistencia.ObtenerClientesBD();
            }
        }

        /// <summary>
        /// Obtiene un objeto ClienteBE a partir del codigo SAP.
        /// </summary>
        /// <param name="codClienteSap"></param>
        /// <returns></returns>
        public ClienteBE ObtenerClienteByCodClienteSap(string codClienteSap)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ClientePersistence persistencia = new ClientePersistence(uow);
                return persistencia.ObtenerClienteByCodClienteSap(codClienteSap);
            }
        }

        #endregion

        #region Guardar

        /// <summary>
        /// Método que guarda un cliente almacenado en memoria con su identificador de cliente
        /// </summary>
        /// <param name="cliente">Cliente a guardar</param>
        /// <returns>ClienteBE guardado</returns>
        public ClienteBE GuardarClienteDeMemoria(ClienteBE cliente, IUnitOfWork uow)
        {
            ClientePersistence persistencia = new ClientePersistence(uow);
            return persistencia.GuardarClienteDeMemoria(cliente);
        }

        /// <summary>
        /// Método que guarda un cliente.
        /// </summary>
        /// <param name="cliente">Cliente a guardar</param>
        /// <param name="uow"></param>
        /// <returns>ClienteBE guardado</returns>
        public ClienteBE GuardarCliente(ClienteBE cliente)
        {
            IUnitOfWork uow = new UnitOfWork();
            ClientePersistence persistencia = new ClientePersistence(uow);
            ClienteBE clienteGuardado = persistencia.GuardarCliente(cliente);
            uow.Save();
            return clienteGuardado;
        }

        /// <summary>
        /// Método que guarda un cliente.
        /// </summary>
        /// <param name="cliente">Cliente a guardar</param>
        /// <param name="codZona">Código de zona</param>
        /// <returns></returns>
        public ClienteBE GuardarCliente(ClienteBE cliente, string codZona)
        {
            IUnitOfWork uow = new UnitOfWork();
            ClientePersistence persistencia = new ClientePersistence(uow);
            ClienteBE clienteGuardado = persistencia.GuardarCliente(cliente, codZona);
            uow.Save();
            return clienteGuardado;
        }

        #endregion

        #endregion
    }
}
