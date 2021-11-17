using Correos.SimuladorOfertas.DTOs;
using Correos.SimuladorOfertas.Persistence;
using Correos.SimuladorOfertas.InOutLight;
using System.Linq;
using System.Collections.ObjectModel;

namespace Correos.SimuladorOfertas.Business
{
    public class EstadoOfertaBL
    {
        #region Métodos Publicos

        /// <summary>
        /// Método que actualiza los estados de las ofertas y sus posiciones
        /// </summary>
        private void ActualizarEstadosOfertasPosiciones(string usuario, string password, Collection<string> listaCodOfertaSAP)
        {
            //Procesar la respuesta.            
            Collection<EstadoPosicionBE> objRespuestaPosiciones = new Collection<EstadoPosicionBE>();
            Collection<EstadoOfertaBE> objRespuestaOfertas = null;
            if (SSOHelper.Instance.LogarConSSO)
            {
                CommunicatorLight conectorSAP = new CommunicatorLight(SSOHelper.Instance.Usuario, password);
                //SSOHelper.Instance.ActualizarCookiePortal();
                SSOHelper.Instance.InicializarWSLight(conectorSAP.Cliente);
                objRespuestaOfertas = conectorSAP.ZCStatusOfertasRfc(listaCodOfertaSAP, out objRespuestaPosiciones);
                SSOHelper.Instance.LimpiarWSLight();
            }
            else
            {
                CommunicatorLight conectorSAP = new CommunicatorLight(usuario, password);
                objRespuestaOfertas = conectorSAP.ZCStatusOfertasRfc(listaCodOfertaSAP, out objRespuestaPosiciones);
            }

            OfertaBL ofertaBL = new OfertaBL();
            ProductoBL productoBL = new ProductoBL();
            ModalidadNegociacionProductoSAPBL modalidadNegociacionBL = new ModalidadNegociacionProductoSAPBL();

            Collection<OfertaBE> listaTodasOfertasPorCodOfertaSAP = ofertaBL.ObtenerOfertas();

            foreach (EstadoPosicionBE item in objRespuestaPosiciones)
            {
                OfertaBE oferta = listaTodasOfertasPorCodOfertaSAP.FirstOrDefault(x => x.CodOfertaSAP.Equals(item.CodOfertaSAP));
                item.IdOferta = oferta.idOferta;

                ProductoBE producto = productoBL.ObtenerProducto(item.CodProductoSAP);
                if (producto != null)
                {
                    item.IdProducto = producto.idProducto;
                    item.IdModalidadNegociacion = modalidadNegociacionBL.ObtenerIdentificadorModalidadNegocioProductoByFK(item.CodProductoSAP, item.ModalidadNegociacion).Value;
                }                
            }

            //Actualizar los estados de ofertas y posiciones
            using (IUnitOfWork uow = new UnitOfWork())
            {
                OfertaPersistence persistencia = new OfertaPersistence(uow);
                persistencia.ActualizarEstadosSAP(objRespuestaOfertas, objRespuestaPosiciones);
                uow.Save();
            }
        }

        /// <summary>
        /// actualiza el estado de las ofertas de sap
        /// </summary>
        /// <param name="usuario"></param>
        /// <param name="password"></param>
        /// <param name="listaCodOfertaSAP"></param>
        public void ActualizarEstadosSAP(string usuario, string password, Collection<string> listaCodOfertaSAP)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                Collection<EstadoOfertaBE> estados = new Collection<EstadoOfertaBE>();
                this.ActualizarEstadosOfertasPosiciones(usuario, password, listaCodOfertaSAP);

                //Se guarda el contexto
                uow.Save();
            }

        }

        #endregion
    }
}
