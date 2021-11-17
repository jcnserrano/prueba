using Correos.SimuladorOfertas.DTOs;
using Correos.SimuladorOfertas.DTOs.Prospecto;
using Correos.SimuladorOfertas.InOutLight;
using Correos.SimuladorOfertas.InOutLight.ServiceWCF_Light;
using Correos.SimuladorOfertas.Persistence;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;

namespace Correos.SimuladorOfertas.Business
{
    public class ProspectoBL
    {
        #region Obtener
        public Collection<EntidadLegalBE> ObtenerListaEntidadLegal()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProspectoPersistence objPersistencia = new ProspectoPersistence(uow);
                return objPersistencia.ObtenerListaEntidadLegal();
            }
        }

        public Collection<CanalAccesoCorreosBE> ObtenerListaCanalAccesoCorreos()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProspectoPersistence objPersistencia = new ProspectoPersistence(uow);
                return objPersistencia.ObtenerListaCanalAccesoCorreos();
            }
        }

        public Collection<CondicionPagoBE> ObtenerListaCondicionPago()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProspectoPersistence objPersistencia = new ProspectoPersistence(uow);
                return objPersistencia.ObtenerListaCondicionPago();
            }
        }

        public Collection<EsquemaSEPABE> ObtenerListaEsquemaSEPA()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProspectoPersistence objPersistencia = new ProspectoPersistence(uow);
                return objPersistencia.ObtenerListaEsquemaSEPA();
            }
        }

        public Collection<FormaPagoBE> ObtenerListaFormaPago()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProspectoPersistence objPersistencia = new ProspectoPersistence(uow);
                return objPersistencia.ObtenerListaFormaPago();
            }
        }

        public Collection<FuncionPersonaContactoBE> ObtenerListaFuncionPersonaContacto()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProspectoPersistence objPersistencia = new ProspectoPersistence(uow);
                return objPersistencia.ObtenerListaFuncionPersonaContacto();
            }
        }

        public Collection<GrupoClientesBE> ObtenerListaGrupoClientes()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProspectoPersistence objPersistencia = new ProspectoPersistence(uow);
                return objPersistencia.ObtenerListaGrupoClientes();
            }
        }

        public Collection<IdiomaBE> ObtenerListaIdioma()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProspectoPersistence objPersistencia = new ProspectoPersistence(uow);
                return objPersistencia.ObtenerListaIdioma();
            }
        }

        public Collection<PaisBE> ObtenerListaPais()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProspectoPersistence objPersistencia = new ProspectoPersistence(uow);
                return objPersistencia.ObtenerListaPais();
            }
        }

        public Collection<RamoBE> ObtenerListaRamo()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProspectoPersistence objPersistencia = new ProspectoPersistence(uow);
                return objPersistencia.ObtenerListaRamo();
            }
        }

        public Collection<SgrandesCuentasBE> ObtenerListaSgrandesCuentas()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProspectoPersistence objPersistencia = new ProspectoPersistence(uow);
                return objPersistencia.ObtenerListaSgrandesCuentas();
            }
        }

        public Collection<TipoContratoBE> ObtenerListaTipoContrato()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProspectoPersistence objPersistencia = new ProspectoPersistence(uow);
                return objPersistencia.ObtenerListaTipoContrato();
            }
        }

        public Collection<TipoRenegociacionBE> ObtenerListaTipoRenegociacion()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProspectoPersistence objPersistencia = new ProspectoPersistence(uow);
                return objPersistencia.ObtenerListaTipoRenegociacion();
            }
        }

        public Collection<TipoViaBE> ObtenerListaTipoVia()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProspectoPersistence objPersistencia = new ProspectoPersistence(uow);
                return objPersistencia.ObtenerListaTipoVia();
            }
        }

        public Collection<TratamientoPersonaContactoBE> ObtenerListaTratamientoPersonaContacto()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProspectoPersistence objPersistencia = new ProspectoPersistence(uow);
                return objPersistencia.ObtenerListaTratamientoPersonaContacto();
            }
        }

        public Collection<ClaseClienteBE> ObtenerListaClaseCliente()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProspectoPersistence objPersistencia = new ProspectoPersistence(uow);
                return objPersistencia.ObtenerListaClaseCliente();
            }
        }

        public string ObtenerZonaProspecto(string region)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProspectoPersistence objPersistencia = new ProspectoPersistence(uow);
                return objPersistencia.ObtenerZonaProspecto(region);
            }
        }

        public Collection<ZonaProspectoBE> ObtenerListaZonasPosibles()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProspectoPersistence objPersistencia = new ProspectoPersistence(uow);
                return objPersistencia.ObtenerListaZonas();
            }
        }
        #endregion

        #region Creación y edición de prospectos
        public ResultadoCargaBE ActualizarProspectoSAP(ProspectoBE prospecto, string usuario, string password)
        {
            ZCProspectosRfcResponse respuesta = null;
            if (SSOHelper.Instance.LogarConSSO)
            {
                CommunicatorLight conectorSAP = new CommunicatorLight(SSOHelper.Instance.Usuario, password);
                //SSOHelper.Instance.ActualizarCookiePortal();
                SSOHelper.Instance.InicializarWSLight(conectorSAP.Cliente);
                respuesta = conectorSAP.ZCCrearProspecto(prospecto, usuario, "2");
                SSOHelper.Instance.LimpiarWSLight();
            }
            else
            {
                CommunicatorLight conectorSAP = new CommunicatorLight(usuario, password);
                respuesta = conectorSAP.ZCCrearProspecto(prospecto, usuario, "2");
            }

            ResultadoCargaBE resultado = new ResultadoCargaBE();
            Collection<ErrorCargaBE> listaErrores = new Collection<ErrorCargaBE>();
            if (respuesta.ItErrores != null)
            {
                foreach (ZepErroresProsGenerad errorSAP in respuesta.ItErrores)
                {
                    listaErrores.Add(new ErrorCargaBE() { error = errorSAP.Message });
                }
            }
            resultado.errores = listaErrores;

            resultado.Cliente = respuesta.WaCargaBpOut.Partner;

            return resultado;
        }

        public ResultadoCargaBE CrearNuevoProspectoSAP(ProspectoBE prospecto, string usuario, string password)
        {
            ZCProspectosRfcResponse respuesta = null;
            if (SSOHelper.Instance.LogarConSSO)
            {
                CommunicatorLight conectorSAP = new CommunicatorLight(SSOHelper.Instance.Usuario, password);
                //SSOHelper.Instance.ActualizarCookiePortal();
                SSOHelper.Instance.InicializarWSLight(conectorSAP.Cliente);
                respuesta = conectorSAP.ZCCrearProspecto(prospecto, usuario, "1");
                SSOHelper.Instance.LimpiarWSLight();
            }
            else
            {
                CommunicatorLight conectorSAP = new CommunicatorLight(usuario, password);
                respuesta = conectorSAP.ZCCrearProspecto(prospecto, usuario, "1");
            }

            ResultadoCargaBE resultado = new ResultadoCargaBE();
            Collection<ErrorCargaBE> listaErrores = new Collection<ErrorCargaBE>();
            if (respuesta.ItErrores != null)
            {
                foreach (ZepErroresProsGenerad errorSAP in respuesta.ItErrores)
                {
                    listaErrores.Add(new ErrorCargaBE() { error = errorSAP.Message });
                }
            }

            resultado.errores = listaErrores;

            resultado.Cliente = respuesta.WaCargaBpOut.Partner;

            return resultado;
        }

        public ResultBE CompruebaNifCex(ProspectoBE prospecto, string usuario, string password, bool esNuevo, string comprobar)
        {
            String nombreCliente = esNuevo ? String.Empty : prospecto.IdCuenta;
            ResultBE resultado;

            if (SSOHelper.Instance.LogarConSSO)
            {
                CommunicatorLight conectorSAP = new CommunicatorLight(SSOHelper.Instance.Usuario, password);
                //SSOHelper.Instance.ActualizarCookiePortal();

                SSOHelper.Instance.InicializarWSLight(conectorSAP.Cliente);
                resultado = conectorSAP.ZCompruebaNifCexRfc(nombreCliente, prospecto.Nif, prospecto.Pais, usuario, comprobar);
                SSOHelper.Instance.LimpiarWSLight();
               
            }
            else
            {
                CommunicatorLight conectorSAP = new CommunicatorLight(usuario, password);
                resultado = conectorSAP.ZCompruebaNifCexRfc(nombreCliente, prospecto.Nif, prospecto.Pais, usuario, comprobar);
            }

            return resultado;
        }
        
        public ResultBE CompruebaNifCex(String cliente, string nif, string pais, string usuario, string password, bool esNuevo, string comprobar)
        {
            //String nombreCliente = esNuevo ? String.Empty : cliente.IdCuenta;
            ResultBE resultado;

            if (SSOHelper.Instance.LogarConSSO)
            {
                CommunicatorLight conectorSAP = new CommunicatorLight(SSOHelper.Instance.Usuario, password);
                //SSOHelper.Instance.ActualizarCookiePortal();

                SSOHelper.Instance.InicializarWSLight(conectorSAP.Cliente);
                resultado = conectorSAP.ZCompruebaNifCexRfc(cliente, nif, pais, usuario, comprobar);
                SSOHelper.Instance.LimpiarWSLight();

            }
            else
            {
                CommunicatorLight conectorSAP = new CommunicatorLight(usuario, password);
                resultado = conectorSAP.ZCompruebaNifCexRfc(cliente, nif, pais, usuario, comprobar);
            }

            return resultado;
        }

        public ProspectoBE ObtenerProspectoPorIdProspectoSAP(string idProspecto, string usuario, string password)
        {
            if (SSOHelper.Instance.LogarConSSO)
            {
                CommunicatorLight conectorSAP = new CommunicatorLight(SSOHelper.Instance.Usuario, password);
                //SSOHelper.Instance.ActualizarCookiePortal();
                SSOHelper.Instance.InicializarWSLight(conectorSAP.Cliente);
                ProspectoBE resultado = conectorSAP.ZCObtenerProspecto(idProspecto, usuario);
                SSOHelper.Instance.LimpiarWSLight();
                return resultado;
            }
            else
            {
                CommunicatorLight conectorSAP = new CommunicatorLight(usuario, password);
                return conectorSAP.ZCObtenerProspecto(idProspecto, usuario);
            }
        }
        #endregion
    }
}
