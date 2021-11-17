using Correos.SimuladorOfertas.Common;
using Correos.SimuladorOfertas.Persistence;
using System.ComponentModel;

namespace Correos.SimuladorOfertas.Business
{
    public class GestionDatosMaestrosBL
    {
        #region Métodos públicos

        /// <summary>
        /// Método que actualiza los datos maestros asociados al usuario
        /// </summary>
        /// <param name="usuario">Identificador del usuario</param>
        /// <param name="password">Contraseña del usuario</param>
        public void ActualizarDatosMaestros(string usuario, string password
                                                    , BackgroundWorker bgw
                                                    , bool esDescargaManual
                                                    , bool debeActualizarDefinicion
                                                    , bool debeActualizarCoeficientePotencialidad
                                                    , bool debeactualizarStatus
                                                    , bool debeActualizarValoresCubicaje
                                                    , bool debeActualizarEstadoSincro
                                                    , bool debeActualizarDefinicionTodos
                                                    , string listaProductosNuevos
                                                    )
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                uow.DeshabilitarAutodeteccionDeCambios();
                if (debeActualizarDefinicion)
                {
                    //Comprobamos los productos que tuvieramos marcados como pendientes.
                    ConfiguracionProductosBL configProductosBL = new ConfiguracionProductosBL();
                    bgw.ReportProgress((int)MensajesBGWEnum.ActualizandoDefinicionesProductos);
                    configProductosBL.ObtenerDefinicionProductosFromSAP(usuario, password, uow, esDescargaManual, debeActualizarDefinicionTodos, listaProductosNuevos);
                }

                if (debeActualizarCoeficientePotencialidad)
                {
                    //Se comprueba si es necesario solicitar a SAP que devuelva la lista de coeficientes para los productos y VA que hay almacenados
                    CoeficienteBL coeficienteBL = new CoeficienteBL();
                    bgw.ReportProgress((int)MensajesBGWEnum.ActualizandoCoeficientesUsuario);
                    coeficienteBL.ComprobarCoeficientesProductos(usuario, password, uow);

                    //Se comprueba si es necesario solicitar a SAP que devuelva la lista de potencialidades para los productos y VA que hay almacenados
                    PotencialidadBL potencialidadBL = new PotencialidadBL();
                    bgw.ReportProgress((int)MensajesBGWEnum.ActualizandoPotencialidadUsuario);
                    potencialidadBL.ComprobarPotencialidadesProductos(usuario, password, uow);
                }

                if (debeactualizarStatus)
                {
                    //Comprobamos el status de todas las ofertas al conectarse.
                    OfertaBL ofertaBL = new OfertaBL();
                    bgw.ReportProgress((int)MensajesBGWEnum.ActualizandoStatusOfertas);
                    ofertaBL.ActualizarStatusOfertasSAP(usuario, password);
                }

                if (debeActualizarValoresCubicaje)
                {
                    // TODO Descomentar para evolutivo de peso volumétrico
                    bgw.ReportProgress((int)MensajesBGWEnum.ActualizandoValoresCubicaje);
                    OfertaBL ofertaBL = new OfertaBL();
                    ofertaBL.ActualizarCubicajeOfertas(usuario, password);
                }

                if (debeActualizarEstadoSincro)
                {
                    //Comprobamos el status de todas las ofertas que se estén sincronizando.
                    OfertaBL ofertaBL = new OfertaBL();
                    bgw.ReportProgress((int)MensajesBGWEnum.ActualizandoEstadoSincronizacionOfertas);
                    ofertaBL.ActualizarEstadosSincronizaciones(usuario, password);                    
                }

                //Se guardan los datos
                uow.Save();
            }
        }

        /// <summary>
        /// Método que actualiza los datos maestros asociados al usuario
        /// </summary>
        /// <param name="usuario">Identificador del usuario</param>
        /// <param name="password">Contraseña del usuario</param>
        public void ActualizarDatosMaestros(string usuario, string password
                                                    , BackgroundWorker bgw
                                                    , bool esDescargaManual
                                                    , bool debeActualizarDefinicion
                                                    , bool debeActualizarCoeficientePotencialidad
                                                    , bool debeactualizarStatus
                                                    , bool debeActualizarValoresCubicaje
                                                    , bool debeactualizarAgrupacionesTipologia
                                                    , ref bool AgrupacionTipologiasActualizadas
                                                    , bool debeActualizarEstadoSincro
                                                    , bool debeActualizarDefinicionTodos
                                                    , string listaProductosNuevos
                                                    )
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                uow.DeshabilitarAutodeteccionDeCambios();
                if (debeActualizarDefinicion)
                {
                    //Comprobamos los productos que tuvieramos marcados como pendientes.
                    ConfiguracionProductosBL configProductosBL = new ConfiguracionProductosBL();

                    //JCNS. TARIFAS 2020
                    //bgw.ReportProgress((int)MensajesBGWEnum.ActualizandoDefinicionesProductos);
                    if (debeActualizarDefinicionTodos == true)
                    {
                        bgw.ReportProgress((int)MensajesBGWEnum.ActualizandoDefinicionesProductosTARIFAS);
                        configProductosBL.ObtenerDefinicionProductosFromSAP_Uno_A_Uno(usuario, password, uow, esDescargaManual, debeActualizarDefinicionTodos);
                    }
                    else
                    {
                        bgw.ReportProgress((int)MensajesBGWEnum.ActualizandoDefinicionesProductos);
                        //JCNS. DESCARGA / ACTUALIZACIÓN PRODUCTOS
                        configProductosBL.ObtenerDefinicionProductosFromSAP(usuario, password, uow, esDescargaManual, debeActualizarDefinicionTodos, listaProductosNuevos);
                        
                        
                    }



                }

                if (debeActualizarCoeficientePotencialidad)
                {
                    //Se comprueba si es necesario solicitar a SAP que devuelva la lista de coeficientes para los productos y VA que hay almacenados
                    CoeficienteBL coeficienteBL = new CoeficienteBL();
                    bgw.ReportProgress((int)MensajesBGWEnum.ActualizandoCoeficientesUsuario);
                    coeficienteBL.ComprobarCoeficientesProductos(usuario, password, uow);

                    //Se comprueba si es necesario solicitar a SAP que devuelva la lista de potencialidades para los productos y VA que hay almacenados
                    PotencialidadBL potencialidadBL = new PotencialidadBL();
                    bgw.ReportProgress((int)MensajesBGWEnum.ActualizandoPotencialidadUsuario);
                    potencialidadBL.ComprobarPotencialidadesProductos(usuario, password, uow);
                }

                if (debeactualizarStatus)
                {
                    //Comprobamos el status de todas las ofertas al conectarse.
                    OfertaBL ofertaBL = new OfertaBL();
                    bgw.ReportProgress((int)MensajesBGWEnum.ActualizandoStatusOfertas);
                    ofertaBL.ActualizarStatusOfertasSAP(usuario, password);
                }

                if (debeActualizarValoresCubicaje)
                {
                    // TODO Descomentar para evolutivo de peso volumétrico
                    bgw.ReportProgress((int)MensajesBGWEnum.ActualizandoValoresCubicaje);
                    OfertaBL ofertaBL = new OfertaBL();
                    ofertaBL.ActualizarCubicajeOfertas(usuario, password);
                }

                if (debeactualizarAgrupacionesTipologia)
                {
                    TipologiaClienteBL tipologiaClienteBL = new TipologiaClienteBL();
                    bgw.ReportProgress((int)MensajesBGWEnum.ActualizandoValoresAgrupacionesTipologia);
                    //Comprobamos si es necesario guardar las agrupaciones de tipologías y si ha sido así avisamos al usuario
                    tipologiaClienteBL.ObtenerAgrupacionesTipologiasDeSAP(usuario, password, System.DateTime.Now, uow, ref AgrupacionTipologiasActualizadas);
                }

                if (debeActualizarEstadoSincro)
                {
                    //Comprobamos el status de todas las ofertas que se estén sincronizando.
                    OfertaBL ofertaBL = new OfertaBL();
                    bgw.ReportProgress((int)MensajesBGWEnum.ActualizandoEstadoSincronizacionOfertas);
                    ofertaBL.ActualizarEstadosSincronizaciones(usuario, password);
                }

                //Se guardan los datos
                uow.Save();

            }
        }

        #endregion
    }
}
