using Correos.SimuladorOfertas.Common;
using Correos.SimuladorOfertas.Common.Enums;
using Correos.SimuladorOfertas.Common.Extensions;
using Correos.SimuladorOfertas.DTOs;
using Correos.SimuladorOfertas.InOutLight;
using Correos.SimuladorOfertas.Persistence;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace Correos.SimuladorOfertas.Business
{
    public class ConfiguracionProductosBL
    {
        private Collection<ProductoBE> ListaProductos { get; set; } 

        #region Métodos Públicos

        #region Descargar CONFIGURACION del producto

        public DescargaOfertaBE ObtenerProductosFromSAP(String codOfertaSAP, string usuario, string password, IUnitOfWork uow)
        {
            DescargaOfertaBE result = new DescargaOfertaBE();

            try
            {
                #region variables usadas

                //lista de variables adaptadas a nuestro modelo de negocio
                result.listaDestinos = new Collection<ConfiguracionDestinoOfertaBE>();
                result.listaTramos = new Collection<ConfiguracionTramoOfertaBE>();
                result.listaPosiciones = new Collection<ConfiguracionPosicionSAPBE>();
                result.listaGruposTramo = new Collection<GrupoTramoBE>();
                result.listaCaracteristicas = new Collection<ConfiguracionCaracteristicaBE>();
                result.listaProductosOferta = new Collection<ProductoOfertaBE>();
                result.listaConfiguracionesValorAnadido = new Collection<ConfiguracionValorAnadidoBE>();
                result.ofertaDescargada = new OfertaBE();

                #endregion

                #region Llamada al servicio
                
                if (SSOHelper.Instance.LogarConSSO)
                {
                    CommunicatorLight conectorSAP = new CommunicatorLight(SSOHelper.Instance.Usuario, password);
                    //SSOHelper.Instance.ActualizarCookiePortal();
                    SSOHelper.Instance.InicializarWSLight(conectorSAP.Cliente);
                    result.resultadoDescarga = conectorSAP.ZCPosicionesASimulaRfc(codOfertaSAP, out result.listaDestinos, out result.listaPosiciones,
                                        out result.listaTramos, out result.listaGruposTramo, out result.ofertaDescargada, out result.listaCaracteristicas, out result.listaConfiguracionesValorAnadido);
                    SSOHelper.Instance.LimpiarWSLight();
                }
                else
                {
                    CommunicatorLight conectorSAP = new CommunicatorLight(usuario, password);
                    result.resultadoDescarga = conectorSAP.ZCPosicionesASimulaRfc(codOfertaSAP, out result.listaDestinos, out result.listaPosiciones,
                                       out result.listaTramos, out result.listaGruposTramo, out result.ofertaDescargada, out result.listaCaracteristicas, out result.listaConfiguracionesValorAnadido);
                }
                                
                #endregion
            }
            catch (Exception e)
            {
                System.Console.WriteLine(e.Message);
                result = null;
            }

            return result;
        }

        /// <summary>
        /// Método que obtiene los datos de configuración para una oferta creada en SAP CRM
        /// </summary>
        /// <param name="oferta">Entidad OfertaBE</param>
        /// <param name="listaActualizadaProductosOferta">Lista de productos que contiene la oferta</param>
        /// <param name="usuario">Usuario</param>
        /// <param name="password">Password</param>
        /// <param name="uow">Objeto transaccional</param>
        /// <returns>Resultado de la operación</returns>
        public ResultBE ObtenerConfiguracionProductosFromSAP(OfertaBE oferta, Collection<ProductoOfertaBE> listaActualizadaProductosOferta, string usuario, string password, IUnitOfWork uow, DescargaOfertaBE resultadoYaDescargado)
        {
            ResultBE result = new ResultBE();

            try
            {
                #region variables usadas

                //lista de variables adaptadas a nuestro modelo de negocio
                Collection<ConfiguracionDestinoOfertaBE> listaDestinos = new Collection<ConfiguracionDestinoOfertaBE>();
                Collection<ConfiguracionTramoOfertaBE> listaTramos = new Collection<ConfiguracionTramoOfertaBE>();
                Collection<ConfiguracionPosicionSAPBE> listaPosiciones = new Collection<ConfiguracionPosicionSAPBE>();
                Collection<GrupoTramoBE> listaGruposTramo = new Collection<GrupoTramoBE>();
                Collection<ConfiguracionCaracteristicaBE> listaCaracteristicas = new Collection<ConfiguracionCaracteristicaBE>();
                Collection<ProductoOfertaBE> listaProductosOferta = new Collection<ProductoOfertaBE>();
                Collection<ConfiguracionValorAnadidoBE> listaConfiguracionesValorAnadido = new Collection<ConfiguracionValorAnadidoBE>();
                OfertaBE ofertaDescargada = new OfertaBE();

                #endregion

                #region Llamada al servicio

                ResultBE objResultado = null;

                if (resultadoYaDescargado == null)
                {
                if (SSOHelper.Instance.LogarConSSO)
                {
                    CommunicatorLight conectorSAP = new CommunicatorLight(SSOHelper.Instance.Usuario, password);
                    //SSOHelper.Instance.ActualizarCookiePortal();
                    SSOHelper.Instance.InicializarWSLight(conectorSAP.Cliente);
                    objResultado = conectorSAP.ZCPosicionesASimulaRfc(oferta.CodOfertaSAP, out listaDestinos, out listaPosiciones,
                                        out listaTramos, out listaGruposTramo, out ofertaDescargada, out listaCaracteristicas, out listaConfiguracionesValorAnadido);
                    SSOHelper.Instance.LimpiarWSLight();
                }
                else
                {
                    CommunicatorLight conectorSAP = new CommunicatorLight(usuario, password);
                    objResultado = conectorSAP.ZCPosicionesASimulaRfc(oferta.CodOfertaSAP, out listaDestinos, out listaPosiciones,
                    out listaTramos, out listaGruposTramo, out ofertaDescargada, out listaCaracteristicas, out listaConfiguracionesValorAnadido);
                }
                }
                else
                {
                    result = resultadoYaDescargado.resultadoDescarga;
                    listaDestinos = resultadoYaDescargado.listaDestinos;
                    listaTramos= resultadoYaDescargado.listaTramos;
                    listaPosiciones = resultadoYaDescargado.listaPosiciones;
                    listaGruposTramo = resultadoYaDescargado.listaGruposTramo;
                    ofertaDescargada = resultadoYaDescargado.ofertaDescargada; 
                    listaCaracteristicas = resultadoYaDescargado.listaCaracteristicas;
                    listaConfiguracionesValorAnadido = resultadoYaDescargado.listaConfiguracionesValorAnadido;                    
                }

                #endregion

                #region Guardar Datos en la base de datos

                //Se vacía la lista almacenada en memoria
                InformacionEstatica.ListaProductosOfertaBE = new Collection<ProductoOfertaBE>();
                //Se guarda la configuración para Posiciones
                PosicionPersistence posicionPersistence = new PosicionPersistence(uow);

                foreach (ProductoOfertaBE productoGuardado in listaActualizadaProductosOferta)
                {
                    posicionPersistence.GuardarConfiguracionPosicion(listaPosiciones.Where(x => x.CodAnexoSAP.Equals(productoGuardado.Anexo) && x.CodProductoSAP.Equals(productoGuardado.CodProductoSAP) && x.CodModalidadNegociacion.Equals(productoGuardado.CodModalidadNegociacion)).ToList<ConfiguracionPosicionSAPBE>().ToCollection<ConfiguracionPosicionSAPBE>());
                }
                  
                //Se guarda la configuración para ValorAñadido
                ValorAnadidoProductoPersistence vaPersistencia = new ValorAnadidoProductoPersistence(uow);
                //vaPersistencia.GuardarConfiguracionValorAnadidoProducto(listaActualizadaProductosOferta, listaPosiciones.Where(x => x.CodAnexoSAP.Equals(po.Anexo) && x.CodProductoSAP.Equals(po.CodProductoSAP) && x.CodModalidadNegociacion.Equals(po.CodModalidadNegociacion)).ToList<ConfiguracionPosicionSAPBE>().ToCollection<ConfiguracionPosicionSAPBE>());
                vaPersistencia.GuardarConfiguracionValorAnadidoProducto(listaActualizadaProductosOferta, listaPosiciones, listaConfiguracionesValorAnadido);

                foreach (ProductoOfertaBE po in InformacionEstatica.ListaProductosOfertaBE)
                {                                  

                    //Se guarda la configuración para Destino
                    DestinoPersistence destinoPersistence = new DestinoPersistence(uow);
                    destinoPersistence.GuardarConfiguracionDestinoOferta(listaDestinos.Where(x => x.CodAnexoSAP.Equals(po.Anexo) && x.CodProductoSAP.Equals(po.CodProductoSAP) && x.Posicion.Equals(po.Posicion)).ToList<ConfiguracionDestinoOfertaBE>().ToCollection<ConfiguracionDestinoOfertaBE>());

                    //Se guarda la configuración para Tramo
                    TramoPersistence tramoPersistence = new TramoPersistence(uow);
                    tramoPersistence.GuardarConfiguracionTramoOferta(listaTramos.Where(x => x.CodAnexoSAP.Equals(po.Anexo) && x.CodProductoSAP.Equals(po.CodProductoSAP) && x.Posicion.Equals(po.Posicion)).ToList<ConfiguracionTramoOfertaBE>().ToCollection<ConfiguracionTramoOfertaBE>());

                    //Se guarda la configuración para ListaPrecios
                    //ListaPreciosPersistence listaPreciosPersistence = new ListaPreciosPersistence(uow);
                    //listaPreciosPersistence.GuardarConfiguracionListaPrecios(listaActualizadaProductosOferta);

                    //Se guarda la lista de grupos de tramo
                    ConfiguracionGruposTramoOfertaPersistence gtPersistence = new ConfiguracionGruposTramoOfertaPersistence(uow);
                    gtPersistence.GuardarConfiguracionGruposTramoOferta(listaGruposTramo.Where(x => x.Anexo.Equals(po.Anexo) && x.CodProductoSAP.Equals(po.CodProductoSAP) && x.Posicion.Equals(po.Posicion)).ToList<GrupoTramoBE>().ToCollection<GrupoTramoBE>());

                    CaracteristicaPersistence caracteristicaPersistence = new CaracteristicaPersistence(uow);
                    foreach (ConfiguracionCaracteristicaBE item in listaCaracteristicas.Where(x => x.CodAnexoSAP.Equals(po.Anexo) && x.CodProductoSAP.Equals(po.CodProductoSAP) && x.Posicion.Equals(po.Posicion)))
                    {
                        caracteristicaPersistence.GuardarConfiguracionDescargarCaracteristicaOferta(item);
                    }
                }


                return result;

                #endregion

            }
            catch (Exception ex)
            {
                result.Resultado = false;
                throw ex;
            }
        }


        #endregion

        #region Descargar DEFINICION de los productos

        /// <summary>
        /// lista de productos a actualizar definiciones
        /// </summary>
        /// <param name="usuario"></param>
        /// <param name="password"></param>
        /// <param name="collection"></param>
        /// <returns></returns>

        // JCNS. AÑADO PARAMETRO bTodos para actualizar los productos sin ActualizacionPendiente.
        public ResultBE ObtenerCollecionDefinicionProductosFromSAP(string usuario, string password, Collection<ProductoBE> collection, bool esDescargaManual, bool bTodos = false, string listaProductosNuevos = "")
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                this.ListaProductos = collection;
                ResultBE objRespuesta = this.ObtenerDefinicionProductosFromSAP(usuario, password, uow, esDescargaManual, bTodos, listaProductosNuevos);
                uow.Save();
                return objRespuesta;
            }
        }

        /// <summary>
        /// Método que realiza la petición a SAP de la configuración para los productos de la oferta
        /// </summary>
        /// <param name="usuario">Usuario que realiza la petición</param>
        /// <param name="password">Contraseña del usuario que realiza la petición</param>
        /// <param name="listaProductos">Lista de los productos cuya definición se quiere guardar</param>
        /// JCNS. AÑADO PARAMETRO bTodos para actualizar los productos sin ActualizacionPendiente.
        /// JCNS. DESCARGA / ACTUALIZACIÓN DE PRODUCTOS (Nuevos o Antiguos)
        public ResultBE ObtenerDefinicionProductosFromSAP(string usuario, string password, IUnitOfWork uow, bool esDescargaManual, bool bTodos, string listaProductosNuevos)
        {
            ProductoPersistence objProductoPersistence = new ProductoPersistence(uow);

            if (!esDescargaManual)
            {
                this.ListaProductos = objProductoPersistence.ObtenerListaProductosDescarga(bTodos);
            }

            ResultBE objRespuesta = new ResultBE();

            if (!string.IsNullOrEmpty(listaProductosNuevos))
            {
                Collection<ProductoBE> listaNuevos = objProductoPersistence.Descarga_Productos_Nuevos(listaProductosNuevos);

                if (listaNuevos.Count > 0)
                {
                    foreach (ProductoBE item in listaNuevos)
                    {
                            this.ListaProductos.Add(new ProductoBE()
                            {
                                CodAnexoSAP = item.CodAnexoSAP,
                                CodProducto = item.CodProducto,
                                ValidezHasta = item.ValidezHasta,
                                ValidezDesde = item.ValidezDesde,
                                ModeloDescuento = item.ModeloDescuento,
                                idProducto = item.idProducto,
                                idAnexoProducto = item.idAnexoProducto,
                                ActualizacionPendiente = item.ActualizacionPendiente
                            });
                    }

                }
            }

            if (this.ListaProductos.Count > 0)
            {
                try
                {


                    #region variables usadas

                    Collection<DescuentoSAPBE> listaDescuentosSAP = new Collection<DescuentoSAPBE>();
                    Collection<ListaPreciosBE> listaPrecios = new Collection<ListaPreciosBE>();
                    Collection<TarifaSAPBE> listaTarifasSAP = new Collection<TarifaSAPBE>();
                    Collection<TipologiaClienteBE> listaTipologiasCliente = new Collection<TipologiaClienteBE>();
                    Collection<PuntosBE> listaPuntos = new Collection<PuntosBE>();
                    Collection<GradosBE> listaGrados = new Collection<GradosBE>();
                    Collection<RegularidadBE> listaRegularidades = new Collection<RegularidadBE>();
                    Collection<PenalizacionRegularidadProductoBE> listaPenalizaciones = new Collection<PenalizacionRegularidadProductoBE>();
                    Collection<RangoPoblacionD2BE> listaRangosPoblacionD2 = new Collection<RangoPoblacionD2BE>();
                    Collection<UmbralBE> listaUmbrales = new Collection<UmbralBE>();
                    Collection<ConfiguracionCaracteristicaBE> listaCaracteristicas = new Collection<ConfiguracionCaracteristicaBE>();
                    Collection<PrecioPorBE> listaPrecioPor = new Collection<PrecioPorBE>();
                    Collection<InternacionalBE> listaProductoInternacional = new Collection<InternacionalBE>();
                    Collection<RelacionProductosBE> listaRelacionProductos = new Collection<RelacionProductosBE>();

                    #endregion

                    #region Llamada al servicio

                    Collection<EsquemaProductoBE> listaEsquemaProducto = null;
                    if (SSOHelper.Instance.LogarConSSO)
                    {
                        CommunicatorLight conectorSAP = new CommunicatorLight(SSOHelper.Instance.Usuario, password);
                        //SSOHelper.Instance.ActualizarCookiePortal();
                        SSOHelper.Instance.InicializarWSLight(conectorSAP.Cliente);
                        listaEsquemaProducto = conectorSAP.ZCConfiguraProductosRfc(this.ListaProductos, out listaDescuentosSAP, out listaPrecios, out listaTarifasSAP,
                                                        out listaTipologiasCliente, out listaPuntos, out listaGrados, out listaRegularidades, out listaPenalizaciones, out listaRangosPoblacionD2, out listaUmbrales, out listaCaracteristicas,
                                                        out listaPrecioPor, out listaProductoInternacional, out listaRelacionProductos);

                        SSOHelper.Instance.LimpiarWSLight();
                    }
                    else
                    {
                        CommunicatorLight conectorSAP = new CommunicatorLight(usuario, password);
                        listaEsquemaProducto = conectorSAP.ZCConfiguraProductosRfc(this.ListaProductos, out listaDescuentosSAP, out listaPrecios, out listaTarifasSAP,
                        out listaTipologiasCliente, out listaPuntos, out listaGrados, out listaRegularidades, out listaPenalizaciones, out listaRangosPoblacionD2, out listaUmbrales, out listaCaracteristicas,
                        out listaPrecioPor, out listaProductoInternacional, out listaRelacionProductos);
                    }
                    #endregion

                    #region Guardar los datos en BBDD

                    //JCNS. TARIFAS 2020. Lo meto en una funcion
                    //// -----------------------------------------------------------------
                    //// Una vez se tienen todos los datos, se guarda en la base de datos:
                    //// -----------------------------------------------------------------

                    ////en listaProductos tenemos la dupla anexo-producto con los índices correspondientes de los que queremos obtener las definiciones
                    //foreach (ProductoBE itemProductoBE in this.ListaProductos)
                    //{
                    //    //JCNS. TARIFAS 2020
                    //    RegistrarAccionesSimulador.GuardarTraza("Actulazando Producto " + itemProductoBE.CodAnexoSAP + " - " + itemProductoBE.CodProducto);

                    //    using (IUnitOfWork nuow = new UnitOfWork())
                    //    {
                    //        nuow.DeshabilitarAutodeteccionDeCambios();

                    //        //Antes de actualizar los distintos productos hay que actualizar la lista de precios                        
                    //        ListaPreciosPersistence persistencia = new ListaPreciosPersistence(nuow);
                    //        persistencia.GuardarListaPrecios(ref listaPrecios);

                    //        ValorAnadidoPersistence vapersistence = new ValorAnadidoPersistence(nuow);
                    //        Collection<ValorAnadidoBE> listaVA = vapersistence.ObtenerValorAnadido();

                    //        //MarcarProducto como obsoleto.
                    //        this.MarcarDefinicionProductoObsoleta(itemProductoBE, nuow);

                    //        //en función del tipo de modelo se guarda su correspondiente definición.
                    //        if (itemProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Costes)))
                    //        {
                    //            //JCNS. DTO-MAX. incluyo listaTipologiasCliente
                    //            //this.GuardarDefinicionProductoCostes(nuow,
                    //            //                                 itemProductoBE,
                    //            //                                 listaEsquemaProducto.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>(),
                    //            //                                 listaDescuentosSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto) && !x.Destino.Equals(string.Empty)).Distinct().ToList<DescuentoSAPBE>().ToCollection<DescuentoSAPBE>(),
                    //            //                                 listaTarifasSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TarifaSAPBE>().ToCollection<TarifaSAPBE>(),
                    //            //                                 listaPuntos.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                    //            //                                 listaGrados.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                    //            //                                 listaPenalizaciones.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PenalizacionRegularidadProductoBE>().ToCollection<PenalizacionRegularidadProductoBE>(),
                    //            //                                 listaRegularidades.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                    //            //                                 ref listaVA,
                    //            //                                 listaProductoInternacional.FirstOrDefault(x => x.Producto.Equals(itemProductoBE.CodProducto) && x.Anexo.Equals(itemProductoBE.CodAnexoSAP) && x.VA.Trim() == ""),
                    //            //                                 listaPrecioPor.Where(x => x.ProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PrecioPorBE>().ToCollection<PrecioPorBE>(),

                    //            //                                 listaRelacionProductos,
                    //            //                                 listaCaracteristicas.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<ConfiguracionCaracteristicaBE>().ToCollection<ConfiguracionCaracteristicaBE>());
                    //            this.GuardarDefinicionProductoCostes(nuow,
                    //                                             itemProductoBE,
                    //                                             listaEsquemaProducto.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>(),
                    //                                             listaDescuentosSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto) && !x.Destino.Equals(string.Empty)).Distinct().ToList<DescuentoSAPBE>().ToCollection<DescuentoSAPBE>(),
                    //                                             listaTarifasSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TarifaSAPBE>().ToCollection<TarifaSAPBE>(),
                    //                                             listaPuntos.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                    //                                             listaGrados.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                    //                                             listaPenalizaciones.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PenalizacionRegularidadProductoBE>().ToCollection<PenalizacionRegularidadProductoBE>(),
                    //                                             listaRegularidades.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                    //                                             ref listaVA,
                    //                                             listaProductoInternacional.FirstOrDefault(x => x.Producto.Equals(itemProductoBE.CodProducto) && x.Anexo.Equals(itemProductoBE.CodAnexoSAP) && x.VA.Trim() == ""),
                    //                                             listaPrecioPor.Where(x => x.ProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PrecioPorBE>().ToCollection<PrecioPorBE>(),

                    //                                             listaRelacionProductos,
                    //                                             listaCaracteristicas.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<ConfiguracionCaracteristicaBE>().ToCollection<ConfiguracionCaracteristicaBE>(),
                    //                                             listaTipologiasCliente.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TipologiaClienteBE>().ToCollection<TipologiaClienteBE>()
                    //                                             );
                    //        }
                    //        else if (itemProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Paqueteria)))
                    //        {
                    //            this.GuardarDefinicionProductoPaqueteria(nuow,
                    //                                                 itemProductoBE,
                    //                                                 listaEsquemaProducto.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>(),
                    //                                                 listaDescuentosSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto) && !x.Destino.Equals(string.Empty)).Distinct().ToList<DescuentoSAPBE>().ToCollection<DescuentoSAPBE>(),
                    //                                                 listaTarifasSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TarifaSAPBE>().ToCollection<TarifaSAPBE>(),
                    //                                                 listaTipologiasCliente.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TipologiaClienteBE>().ToCollection<TipologiaClienteBE>(),
                    //                                                 ref listaVA,
                    //                                                 listaProductoInternacional.FirstOrDefault(x => x.Producto.Equals(itemProductoBE.CodProducto) && x.Anexo.Equals(itemProductoBE.CodAnexoSAP) && x.VA.Trim() == ""),
                    //                                                 listaPrecioPor.Where(x => x.ProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PrecioPorBE>().ToCollection<PrecioPorBE>(),
                    //                                                 listaRelacionProductos);
                    //        }
                    //        else if (itemProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Tramos)))
                    //        {
                    //            this.GuardarDefinicionProductoTramos(nuow,
                    //                                             itemProductoBE,
                    //                                             listaEsquemaProducto.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>(),
                    //                                             listaDescuentosSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto) && !x.Destino.Equals(string.Empty)).Distinct().ToList<DescuentoSAPBE>().ToCollection<DescuentoSAPBE>(),
                    //                                             listaTarifasSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TarifaSAPBE>().ToCollection<TarifaSAPBE>(),
                    //                                             listaTipologiasCliente.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TipologiaClienteBE>().ToCollection<TipologiaClienteBE>(),
                    //                                             listaCaracteristicas.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<ConfiguracionCaracteristicaBE>().ToCollection<ConfiguracionCaracteristicaBE>(),
                    //                                             ref listaVA,
                    //                                             listaPrecioPor.Where(x => x.ProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PrecioPorBE>().ToCollection<PrecioPorBE>(),
                    //                                             listaProductoInternacional.FirstOrDefault(x => x.Producto.Equals(itemProductoBE.CodProducto) && x.Anexo.Equals(itemProductoBE.CodAnexoSAP) && x.VA.Trim() == ""),
                    //                                             listaRelacionProductos);
                    //        }
                    //        else if (itemProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Volumetrico)))
                    //        {
                    //            //JCNS. DTO-MAX. incluyo listaTipologiasCliente
                    //            //this.GuardarDefinicionProductoVolumetrico(nuow,
                    //            //                                      itemProductoBE,
                    //            //                                      listaEsquemaProducto.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>(),
                    //            //                                      listaDescuentosSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto) && !x.Destino.Equals(string.Empty)).Distinct().ToList<DescuentoSAPBE>().ToCollection<DescuentoSAPBE>(),
                    //            //                                      listaTarifasSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TarifaSAPBE>().ToCollection<TarifaSAPBE>(),
                    //            //                                      listaGrados.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                    //            //                                      listaPenalizaciones.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PenalizacionRegularidadProductoBE>().ToCollection<PenalizacionRegularidadProductoBE>(),
                    //            //                                      listaRegularidades.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                    //            //                                      listaUmbrales.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                    //            //                                      listaRangosPoblacionD2.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<RangoPoblacionD2BE>().ToCollection<RangoPoblacionD2BE>(),
                    //            //                                      ref listaVA,
                    //            //                                      listaProductoInternacional.FirstOrDefault(x => x.Producto.Equals(itemProductoBE.CodProducto) && x.Anexo.Equals(itemProductoBE.CodAnexoSAP) && x.VA.Trim() == ""),
                    //            //                                      listaPrecioPor.Where(x => x.ProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PrecioPorBE>().ToCollection<PrecioPorBE>(),
                    //            //                                      listaRelacionProductos);
                    //            this.GuardarDefinicionProductoVolumetrico(nuow,
                    //                                                  itemProductoBE,
                    //                                                  listaEsquemaProducto.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>(),
                    //                                                  listaDescuentosSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto) && !x.Destino.Equals(string.Empty)).Distinct().ToList<DescuentoSAPBE>().ToCollection<DescuentoSAPBE>(),
                    //                                                  listaTarifasSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TarifaSAPBE>().ToCollection<TarifaSAPBE>(),
                    //                                                  listaGrados.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                    //                                                  listaPenalizaciones.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PenalizacionRegularidadProductoBE>().ToCollection<PenalizacionRegularidadProductoBE>(),
                    //                                                  listaRegularidades.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                    //                                                  listaUmbrales.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                    //                                                  listaRangosPoblacionD2.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<RangoPoblacionD2BE>().ToCollection<RangoPoblacionD2BE>(),
                    //                                                  ref listaVA,
                    //                                                  listaProductoInternacional.FirstOrDefault(x => x.Producto.Equals(itemProductoBE.CodProducto) && x.Anexo.Equals(itemProductoBE.CodAnexoSAP) && x.VA.Trim() == ""),
                    //                                                  listaPrecioPor.Where(x => x.ProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PrecioPorBE>().ToCollection<PrecioPorBE>(),
                    //                                                  listaRelacionProductos,
                    //                                                  listaTipologiasCliente.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TipologiaClienteBE>().ToCollection<TipologiaClienteBE>()
                    //                                                    );


                    //        }
                    //        else if (itemProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Publicorreo)))
                    //        {
                    //            //JCNS. DTO-MAX. incluyo listaTipologiasCliente
                    //            //this.GuardarDefinicionProductoPublicorreo(nuow,
                    //            //                                      itemProductoBE,
                    //            //                                      listaEsquemaProducto.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>(),
                    //            //                                      listaDescuentosSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto) ).Distinct().ToList<DescuentoSAPBE>().ToCollection<DescuentoSAPBE>(),
                    //            //                                      listaTarifasSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TarifaSAPBE>().ToCollection<TarifaSAPBE>(),
                    //            //                                      listaPuntos.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                    //            //                                      listaGrados.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                    //            //                                      listaPenalizaciones.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PenalizacionRegularidadProductoBE>().ToCollection<PenalizacionRegularidadProductoBE>(),
                    //            //                                      listaRegularidades.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                    //            //                                      ref listaVA,
                    //            //                                      listaProductoInternacional.FirstOrDefault(x => x.Producto.Equals(itemProductoBE.CodProducto) && x.Anexo.Equals(itemProductoBE.CodAnexoSAP) && x.VA.Trim() == ""),
                    //            //                                      listaPrecioPor.Where(x => x.ProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PrecioPorBE>().ToCollection<PrecioPorBE>(),

                    //            //                                      listaRelacionProductos,
                    //            //                                      listaCaracteristicas.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<ConfiguracionCaracteristicaBE>().ToCollection<ConfiguracionCaracteristicaBE>());
                    //            this.GuardarDefinicionProductoPublicorreo(nuow,
                    //                                                  itemProductoBE,
                    //                                                  listaEsquemaProducto.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>(),
                    //                                                  listaDescuentosSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto)).Distinct().ToList<DescuentoSAPBE>().ToCollection<DescuentoSAPBE>(),
                    //                                                  listaTarifasSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TarifaSAPBE>().ToCollection<TarifaSAPBE>(),
                    //                                                  listaPuntos.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                    //                                                  listaGrados.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                    //                                                  listaPenalizaciones.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PenalizacionRegularidadProductoBE>().ToCollection<PenalizacionRegularidadProductoBE>(),
                    //                                                  listaRegularidades.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                    //                                                  ref listaVA,
                    //                                                  listaProductoInternacional.FirstOrDefault(x => x.Producto.Equals(itemProductoBE.CodProducto) && x.Anexo.Equals(itemProductoBE.CodAnexoSAP) && x.VA.Trim() == ""),
                    //                                                  listaPrecioPor.Where(x => x.ProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PrecioPorBE>().ToCollection<PrecioPorBE>(),

                    //                                                  listaRelacionProductos,
                    //                                                  listaCaracteristicas.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<ConfiguracionCaracteristicaBE>().ToCollection<ConfiguracionCaracteristicaBE>(),
                    //                                                  listaTipologiasCliente.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TipologiaClienteBE>().ToCollection<TipologiaClienteBE>()
                    //                                                  );
                    //        }
                    //        else if (itemProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Etiquetas)))
                    //        {
                    //            //JCNS. DTO-MAX. incluyo listaTipologiasCliente
                    //            //this.GuardarDefinicionProductoEtiquetas(nuow,
                    //            //                                    itemProductoBE,
                    //            //                                    listaEsquemaProducto.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>(),
                    //            //                                    listaDescuentosSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto) && !x.Destino.Equals(string.Empty)).Distinct().ToList<DescuentoSAPBE>().ToCollection<DescuentoSAPBE>(),
                    //            //                                    listaTarifasSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TarifaSAPBE>().ToCollection<TarifaSAPBE>(),
                    //            //                                    listaCaracteristicas.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<ConfiguracionCaracteristicaBE>().ToCollection<ConfiguracionCaracteristicaBE>(),
                    //            //                                    ref listaVA,
                    //            //                                    listaProductoInternacional.FirstOrDefault(x => x.Producto.Equals(itemProductoBE.CodProducto) && x.Anexo.Equals(itemProductoBE.CodAnexoSAP) && x.VA.Trim() == ""),
                    //            //                                    listaPrecioPor.Where(x => x.ProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PrecioPorBE>().ToCollection<PrecioPorBE>(),

                    //            //                                    listaRelacionProductos);
                    //            this.GuardarDefinicionProductoEtiquetas(nuow,
                    //                                                itemProductoBE,
                    //                                                listaEsquemaProducto.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>(),
                    //                                                listaDescuentosSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto) && !x.Destino.Equals(string.Empty)).Distinct().ToList<DescuentoSAPBE>().ToCollection<DescuentoSAPBE>(),
                    //                                                listaTarifasSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TarifaSAPBE>().ToCollection<TarifaSAPBE>(),
                    //                                                listaCaracteristicas.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<ConfiguracionCaracteristicaBE>().ToCollection<ConfiguracionCaracteristicaBE>(),
                    //                                                ref listaVA,
                    //                                                listaProductoInternacional.FirstOrDefault(x => x.Producto.Equals(itemProductoBE.CodProducto) && x.Anexo.Equals(itemProductoBE.CodAnexoSAP) && x.VA.Trim() == ""),
                    //                                                listaPrecioPor.Where(x => x.ProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PrecioPorBE>().ToCollection<PrecioPorBE>(),

                    //                                                listaRelacionProductos,
                    //                                                  listaTipologiasCliente.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TipologiaClienteBE>().ToCollection<TipologiaClienteBE>()
                    //                                                );
                    //        }
                    //        else if (itemProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Libros)))
                    //        {
                    //            //JCNS. DTO-MAX. incluyo listaTipologiasCliente
                    //            //this.GuardarDefinicionProductoLibros(nuow,
                    //            //                                 itemProductoBE,
                    //            //                                 listaEsquemaProducto.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>(),
                    //            //                                 listaDescuentosSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto) && !x.Destino.Equals(string.Empty)).Distinct().ToList<DescuentoSAPBE>().ToCollection<DescuentoSAPBE>(),
                    //            //                                 listaTarifasSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TarifaSAPBE>().ToCollection<TarifaSAPBE>(),
                    //            //                                 listaPuntos.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                    //            //                                 listaGrados.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                    //            //                                 listaPenalizaciones.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PenalizacionRegularidadProductoBE>().ToCollection<PenalizacionRegularidadProductoBE>(),
                    //            //                                 listaRegularidades.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                    //            //                                 ref listaVA,
                    //            //                                 listaProductoInternacional.FirstOrDefault(x => x.Producto.Equals(itemProductoBE.CodProducto) && x.Anexo.Equals(itemProductoBE.CodAnexoSAP) && x.VA.Trim() == ""),
                    //            //                                 listaPrecioPor.Where(x => x.ProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PrecioPorBE>().ToCollection<PrecioPorBE>(),


                    //            //                                 listaRelacionProductos,
                    //            //                                 listaCaracteristicas.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<ConfiguracionCaracteristicaBE>().ToCollection<ConfiguracionCaracteristicaBE>());
                    //            this.GuardarDefinicionProductoLibros(nuow,
                    //                                             itemProductoBE,
                    //                                             listaEsquemaProducto.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>(),
                    //                                             listaDescuentosSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto) && !x.Destino.Equals(string.Empty)).Distinct().ToList<DescuentoSAPBE>().ToCollection<DescuentoSAPBE>(),
                    //                                             listaTarifasSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TarifaSAPBE>().ToCollection<TarifaSAPBE>(),
                    //                                             listaPuntos.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                    //                                             listaGrados.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                    //                                             listaPenalizaciones.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PenalizacionRegularidadProductoBE>().ToCollection<PenalizacionRegularidadProductoBE>(),
                    //                                             listaRegularidades.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                    //                                             ref listaVA,
                    //                                             listaProductoInternacional.FirstOrDefault(x => x.Producto.Equals(itemProductoBE.CodProducto) && x.Anexo.Equals(itemProductoBE.CodAnexoSAP) && x.VA.Trim() == ""),
                    //                                             listaPrecioPor.Where(x => x.ProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PrecioPorBE>().ToCollection<PrecioPorBE>(),


                    //                                             listaRelacionProductos,
                    //                                             listaCaracteristicas.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<ConfiguracionCaracteristicaBE>().ToCollection<ConfiguracionCaracteristicaBE>(),
                    //                                                  listaTipologiasCliente.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TipologiaClienteBE>().ToCollection<TipologiaClienteBE>()
                    //                                             );
                    //        }
                    //        else if (itemProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Publicaciones)))
                    //        {
                    //            //JCNS. DTO-MAX. incluyo listaTipologiasCliente
                    //            //this.GuardarDefinicionProductoPublicaciones(nuow,
                    //            //                                        itemProductoBE,
                    //            //                                        listaEsquemaProducto.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>(),
                    //            //                                        listaDescuentosSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto) && !x.Destino.Equals(string.Empty)).Distinct().ToList<DescuentoSAPBE>().ToCollection<DescuentoSAPBE>(),
                    //            //                                        listaTarifasSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TarifaSAPBE>().ToCollection<TarifaSAPBE>(),
                    //            //                                        listaPuntos.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                    //            //                                        listaGrados.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                    //            //                                        listaPenalizaciones.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PenalizacionRegularidadProductoBE>().ToCollection<PenalizacionRegularidadProductoBE>(),
                    //            //                                        listaRegularidades.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                    //            //                                        ref listaVA,
                    //            //                                        listaProductoInternacional.FirstOrDefault(x => x.Producto.Equals(itemProductoBE.CodProducto) && x.Anexo.Equals(itemProductoBE.CodAnexoSAP) && x.VA.Trim() == ""),
                    //            //                                        listaPrecioPor.Where(x => x.ProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PrecioPorBE>().ToCollection<PrecioPorBE>(),
                    //            //                                        listaRelacionProductos,
                    //            //                                        listaCaracteristicas.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<ConfiguracionCaracteristicaBE>().ToCollection<ConfiguracionCaracteristicaBE>());
                    //            this.GuardarDefinicionProductoPublicaciones(nuow,
                    //                                                    itemProductoBE,
                    //                                                    listaEsquemaProducto.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>(),
                    //                                                    listaDescuentosSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto) && !x.Destino.Equals(string.Empty)).Distinct().ToList<DescuentoSAPBE>().ToCollection<DescuentoSAPBE>(),
                    //                                                    listaTarifasSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TarifaSAPBE>().ToCollection<TarifaSAPBE>(),
                    //                                                    listaPuntos.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                    //                                                    listaGrados.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                    //                                                    listaPenalizaciones.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PenalizacionRegularidadProductoBE>().ToCollection<PenalizacionRegularidadProductoBE>(),
                    //                                                    listaRegularidades.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                    //                                                    ref listaVA,
                    //                                                    listaProductoInternacional.FirstOrDefault(x => x.Producto.Equals(itemProductoBE.CodProducto) && x.Anexo.Equals(itemProductoBE.CodAnexoSAP) && x.VA.Trim() == ""),
                    //                                                    listaPrecioPor.Where(x => x.ProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PrecioPorBE>().ToCollection<PrecioPorBE>(),
                    //                                                    listaRelacionProductos,
                    //                                                    listaCaracteristicas.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<ConfiguracionCaracteristicaBE>().ToCollection<ConfiguracionCaracteristicaBE>(),
                    //                                                  listaTipologiasCliente.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TipologiaClienteBE>().ToCollection<TipologiaClienteBE>()
                    //                                                    );
                    //        }


                    //        this.EliminarCoeficientesProducto(itemProductoBE.idProducto, usuario, nuow);

                    //        //Se eliminan las potencialidades del producto antiguo y de los valores añadidos del producto antiguo
                    //        this.EliminarPotencialidadesProducto(itemProductoBE.idProducto, usuario, nuow);

                    //        //Se salva el contexto global de la descarga
                    //        nuow.Save();
                    //    }
                    //}


                    ObtenerDefinicionProductosFromSAP_Actualizar(usuario
                                                                , this.ListaProductos
                                                                , listaEsquemaProducto
                                                                , listaDescuentosSAP
                                                                , listaPrecios
                                                                , listaTarifasSAP
                                                                , listaTipologiasCliente
                                                                , listaPuntos
                                                                , listaGrados
                                                                , listaRegularidades
                                                                , listaPenalizaciones
                                                                , listaRangosPoblacionD2
                                                                , listaUmbrales
                                                                , listaCaracteristicas
                                                                , listaPrecioPor
                                                                , listaProductoInternacional
                                                                , listaRelacionProductos);


                    //JCNS. 
                    RegistrarAccionesSimulador.GuardarTraza("FIN Actualizacion de Productos -*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*");

                    #endregion

                }
                catch (Exception ex)
                {
                    objRespuesta.Resultado = false;

                }

            }
            return objRespuesta;
        }




        /// JCNS. AÑADO PARAMETRO bTodos para actualizar los productos sin ActualizacionPendiente.
        public ResultBE ObtenerDefinicionProductosFromSAP_Uno_A_Uno(string usuario, string password, IUnitOfWork uow, bool esDescargaManual, bool bTodos)
        {
            Collection<ProductoBE> lstProductos = new Collection<ProductoBE>();

            ProductoPersistence objProductoPersistence = new ProductoPersistence(uow);

            if (!esDescargaManual)
            {
                this.ListaProductos = objProductoPersistence.ObtenerListaProductosDescarga(bTodos);
            }

            ResultBE objRespuesta = new ResultBE();

            if (this.ListaProductos.Count == 0) return objRespuesta;

            try
            {
                #region variables usadas

                Collection<DescuentoSAPBE> listaDescuentosSAP = new Collection<DescuentoSAPBE>();
                Collection<ListaPreciosBE> listaPrecios = new Collection<ListaPreciosBE>();
                Collection<TarifaSAPBE> listaTarifasSAP = new Collection<TarifaSAPBE>();
                Collection<TipologiaClienteBE> listaTipologiasCliente = new Collection<TipologiaClienteBE>();
                Collection<PuntosBE> listaPuntos = new Collection<PuntosBE>();
                Collection<GradosBE> listaGrados = new Collection<GradosBE>();
                Collection<RegularidadBE> listaRegularidades = new Collection<RegularidadBE>();
                Collection<PenalizacionRegularidadProductoBE> listaPenalizaciones = new Collection<PenalizacionRegularidadProductoBE>();
                Collection<RangoPoblacionD2BE> listaRangosPoblacionD2 = new Collection<RangoPoblacionD2BE>();
                Collection<UmbralBE> listaUmbrales = new Collection<UmbralBE>();
                Collection<ConfiguracionCaracteristicaBE> listaCaracteristicas = new Collection<ConfiguracionCaracteristicaBE>();
                Collection<PrecioPorBE> listaPrecioPor = new Collection<PrecioPorBE>();
                Collection<InternacionalBE> listaProductoInternacional = new Collection<InternacionalBE>();
                Collection<RelacionProductosBE> listaRelacionProductos = new Collection<RelacionProductosBE>();

                #endregion

                //for Collection<ProductoBE>
                RegistrarAccionesSimulador.GuardarTraza("INICIO Actualizacion de Tarifas -*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*");

                foreach (ProductoBE itemProductoBE in this.ListaProductos)
                {
                    Collection<ProductoBE> auxListaProductos = new Collection<ProductoBE>();
                    auxListaProductos.Add(new ProductoBE()
                    {
                        CodAnexoSAP = itemProductoBE.CodAnexoSAP,
                        CodProducto = itemProductoBE.CodProducto,
                        ValidezHasta = itemProductoBE.ValidezHasta,
                        ValidezDesde = itemProductoBE.ValidezDesde,
                        ModeloDescuento = itemProductoBE.ModeloDescuento,
                        idProducto = itemProductoBE.idProducto,
                        idAnexoProducto = itemProductoBE.idAnexoProducto,
                        ActualizacionPendiente = itemProductoBE.ActualizacionPendiente
                    });



                    #region Llamada al servicio

                    Collection<EsquemaProductoBE> listaEsquemaProducto = null;
                    if (SSOHelper.Instance.LogarConSSO)
                    {
                        CommunicatorLight conectorSAP = new CommunicatorLight(SSOHelper.Instance.Usuario, password);
                        //SSOHelper.Instance.ActualizarCookiePortal();
                        SSOHelper.Instance.InicializarWSLight(conectorSAP.Cliente);
                        listaEsquemaProducto = conectorSAP.ZCConfiguraProductosRfc(auxListaProductos, out listaDescuentosSAP, out listaPrecios, out listaTarifasSAP,
                                                        out listaTipologiasCliente, out listaPuntos, out listaGrados, out listaRegularidades, out listaPenalizaciones, out listaRangosPoblacionD2, out listaUmbrales, out listaCaracteristicas,
                                                        out listaPrecioPor, out listaProductoInternacional, out listaRelacionProductos);

                        SSOHelper.Instance.LimpiarWSLight();
                    }
                    else
                    {
                        CommunicatorLight conectorSAP = new CommunicatorLight(usuario, password);
                        listaEsquemaProducto = conectorSAP.ZCConfiguraProductosRfc(auxListaProductos, out listaDescuentosSAP, out listaPrecios, out listaTarifasSAP,
                        out listaTipologiasCliente, out listaPuntos, out listaGrados, out listaRegularidades, out listaPenalizaciones, out listaRangosPoblacionD2, out listaUmbrales, out listaCaracteristicas,
                        out listaPrecioPor, out listaProductoInternacional, out listaRelacionProductos);
                    }
                    #endregion


                    #region Guardar los datos en BBDD

                    //JCNS. TARIFAS 2020

                    ObtenerDefinicionProductosFromSAP_Actualizar(usuario
                                                                , auxListaProductos
                                                                , listaEsquemaProducto
                                                                , listaDescuentosSAP
                                                                , listaPrecios
                                                                , listaTarifasSAP
                                                                , listaTipologiasCliente
                                                                , listaPuntos
                                                                , listaGrados
                                                                , listaRegularidades
                                                                , listaPenalizaciones
                                                                , listaRangosPoblacionD2
                                                                , listaUmbrales
                                                                , listaCaracteristicas
                                                                , listaPrecioPor
                                                                , listaProductoInternacional
                                                                , listaRelacionProductos);
                    #endregion
                }

                //JCNS. TARIFAS 2020
                RegistrarAccionesSimulador.GuardarTraza("FIN Actualizacion de Tarifas -*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*");

                RegistrarAccionesSimulador.GuardarTraza(new AuditoriaDataTrazaBE("Actualizando script 78_5", "", "", ""));

                VersionesDataBasePersistence cambioVersion = new VersionesDataBasePersistence(uow);
                cambioVersion.actualizar_version_BDVersion(string.Format(Utils.GetValorFromAppConfig(AppSettingsEnum.RutaScripts), AppDomain.CurrentDomain.BaseDirectory, "update78_5.sql"));


            }
            catch (Exception ex)
            {
                objRespuesta.Resultado = false;

            }
            return objRespuesta;
        }




        public void ObtenerDefinicionProductosFromSAP_Actualizar(string usuario
                                                                    , Collection<ProductoBE> listaProductos
                                                                    , Collection<EsquemaProductoBE> listaEsquemaProducto
                                                                    , Collection<DescuentoSAPBE> listaDescuentosSAP
                                                                    , Collection<ListaPreciosBE> listaPrecios
                                                                    , Collection<TarifaSAPBE> listaTarifasSAP
                                                                    , Collection<TipologiaClienteBE> listaTipologiasCliente
                                                                    , Collection<PuntosBE> listaPuntos
                                                                    , Collection<GradosBE> listaGrados
                                                                    , Collection<RegularidadBE> listaRegularidades
                                                                    , Collection<PenalizacionRegularidadProductoBE> listaPenalizaciones
                                                                    , Collection<RangoPoblacionD2BE> listaRangosPoblacionD2
                                                                    , Collection<UmbralBE> listaUmbrales
                                                                    , Collection<ConfiguracionCaracteristicaBE> listaCaracteristicas
                                                                    , Collection<PrecioPorBE> listaPrecioPor
                                                                    , Collection<InternacionalBE> listaProductoInternacional
                                                                    , Collection<RelacionProductosBE> listaRelacionProductos)
        {
            foreach (ProductoBE itemProductoBE in listaProductos)
            {
                //JCNS. TARIFAS 2020
                RegistrarAccionesSimulador.GuardarTraza("INICIO Actualizar Producto " + itemProductoBE.CodAnexoSAP + " - " + itemProductoBE.CodProducto);

                using (IUnitOfWork nuow = new UnitOfWork())
                {
                    nuow.DeshabilitarAutodeteccionDeCambios();

                    //Antes de actualizar los distintos productos hay que actualizar la lista de precios                        
                    ListaPreciosPersistence persistencia = new ListaPreciosPersistence(nuow);
                    persistencia.GuardarListaPrecios(ref listaPrecios);

                    ValorAnadidoPersistence vapersistence = new ValorAnadidoPersistence(nuow);
                    Collection<ValorAnadidoBE> listaVA = vapersistence.ObtenerValorAnadido();

                    //MarcarProducto como obsoleto.
                    this.MarcarDefinicionProductoObsoleta(itemProductoBE, nuow);

                    //en función del tipo de modelo se guarda su correspondiente definición.
                    if (itemProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Costes)))
                    {
                        //JCNS. DTO-MAX. incluyo listaTipologiasCliente
                        //this.GuardarDefinicionProductoCostes(nuow,
                        //                                 itemProductoBE,
                        //                                 listaEsquemaProducto.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>(),
                        //                                 listaDescuentosSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto) && !x.Destino.Equals(string.Empty)).Distinct().ToList<DescuentoSAPBE>().ToCollection<DescuentoSAPBE>(),
                        //                                 listaTarifasSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TarifaSAPBE>().ToCollection<TarifaSAPBE>(),
                        //                                 listaPuntos.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                        //                                 listaGrados.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                        //                                 listaPenalizaciones.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PenalizacionRegularidadProductoBE>().ToCollection<PenalizacionRegularidadProductoBE>(),
                        //                                 listaRegularidades.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                        //                                 ref listaVA,
                        //                                 listaProductoInternacional.FirstOrDefault(x => x.Producto.Equals(itemProductoBE.CodProducto) && x.Anexo.Equals(itemProductoBE.CodAnexoSAP) && x.VA.Trim() == ""),
                        //                                 listaPrecioPor.Where(x => x.ProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PrecioPorBE>().ToCollection<PrecioPorBE>(),

                        //                                 listaRelacionProductos,
                        //                                 listaCaracteristicas.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<ConfiguracionCaracteristicaBE>().ToCollection<ConfiguracionCaracteristicaBE>());
                        this.GuardarDefinicionProductoCostes(nuow,
                                                         itemProductoBE,
                                                         listaEsquemaProducto.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>(),
                                                         listaDescuentosSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto) && !x.Destino.Equals(string.Empty)).Distinct().ToList<DescuentoSAPBE>().ToCollection<DescuentoSAPBE>(),
                                                         listaTarifasSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TarifaSAPBE>().ToCollection<TarifaSAPBE>(),
                                                         listaPuntos.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                                                         listaGrados.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                                                         listaPenalizaciones.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PenalizacionRegularidadProductoBE>().ToCollection<PenalizacionRegularidadProductoBE>(),
                                                         listaRegularidades.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                                                         ref listaVA,
                                                         listaProductoInternacional.FirstOrDefault(x => x.Producto.Equals(itemProductoBE.CodProducto) && x.Anexo.Equals(itemProductoBE.CodAnexoSAP) && x.VA.Trim() == ""),
                                                         listaPrecioPor.Where(x => x.ProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PrecioPorBE>().ToCollection<PrecioPorBE>(),

                                                         listaRelacionProductos,
                                                         listaCaracteristicas.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<ConfiguracionCaracteristicaBE>().ToCollection<ConfiguracionCaracteristicaBE>(),
                                                         listaTipologiasCliente.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TipologiaClienteBE>().ToCollection<TipologiaClienteBE>()
                                                         );
                    }
                    else if (itemProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Paqueteria)))
                    {
                        this.GuardarDefinicionProductoPaqueteria(nuow,
                                                             itemProductoBE,
                                                             listaEsquemaProducto.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>(),
                                                             listaDescuentosSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto) && !x.Destino.Equals(string.Empty)).Distinct().ToList<DescuentoSAPBE>().ToCollection<DescuentoSAPBE>(),
                                                             listaTarifasSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TarifaSAPBE>().ToCollection<TarifaSAPBE>(),
                                                             listaTipologiasCliente.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TipologiaClienteBE>().ToCollection<TipologiaClienteBE>(),
                                                             ref listaVA,
                                                             listaProductoInternacional.FirstOrDefault(x => x.Producto.Equals(itemProductoBE.CodProducto) && x.Anexo.Equals(itemProductoBE.CodAnexoSAP) && x.VA.Trim() == ""),
                                                             listaPrecioPor.Where(x => x.ProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PrecioPorBE>().ToCollection<PrecioPorBE>(),
                                                             listaRelacionProductos);
                    }
                    else if (itemProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Tramos)))
                    {
                        this.GuardarDefinicionProductoTramos(nuow,
                                                         itemProductoBE,
                                                         listaEsquemaProducto.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>(),
                                                         listaDescuentosSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto) && !x.Destino.Equals(string.Empty)).Distinct().ToList<DescuentoSAPBE>().ToCollection<DescuentoSAPBE>(),
                                                         listaTarifasSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TarifaSAPBE>().ToCollection<TarifaSAPBE>(),
                                                         listaTipologiasCliente.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TipologiaClienteBE>().ToCollection<TipologiaClienteBE>(),
                                                         listaCaracteristicas.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<ConfiguracionCaracteristicaBE>().ToCollection<ConfiguracionCaracteristicaBE>(),
                                                         ref listaVA,
                                                         listaPrecioPor.Where(x => x.ProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PrecioPorBE>().ToCollection<PrecioPorBE>(),
                                                         listaProductoInternacional.FirstOrDefault(x => x.Producto.Equals(itemProductoBE.CodProducto) && x.Anexo.Equals(itemProductoBE.CodAnexoSAP) && x.VA.Trim() == ""),
                                                         listaRelacionProductos);
                    }
                    else if (itemProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Volumetrico)))
                    {
                        //JCNS. DTO-MAX. incluyo listaTipologiasCliente
                        //this.GuardarDefinicionProductoVolumetrico(nuow,
                        //                                      itemProductoBE,
                        //                                      listaEsquemaProducto.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>(),
                        //                                      listaDescuentosSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto) && !x.Destino.Equals(string.Empty)).Distinct().ToList<DescuentoSAPBE>().ToCollection<DescuentoSAPBE>(),
                        //                                      listaTarifasSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TarifaSAPBE>().ToCollection<TarifaSAPBE>(),
                        //                                      listaGrados.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                        //                                      listaPenalizaciones.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PenalizacionRegularidadProductoBE>().ToCollection<PenalizacionRegularidadProductoBE>(),
                        //                                      listaRegularidades.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                        //                                      listaUmbrales.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                        //                                      listaRangosPoblacionD2.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<RangoPoblacionD2BE>().ToCollection<RangoPoblacionD2BE>(),
                        //                                      ref listaVA,
                        //                                      listaProductoInternacional.FirstOrDefault(x => x.Producto.Equals(itemProductoBE.CodProducto) && x.Anexo.Equals(itemProductoBE.CodAnexoSAP) && x.VA.Trim() == ""),
                        //                                      listaPrecioPor.Where(x => x.ProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PrecioPorBE>().ToCollection<PrecioPorBE>(),
                        //                                      listaRelacionProductos);
                        this.GuardarDefinicionProductoVolumetrico(nuow,
                                                              itemProductoBE,
                                                              listaEsquemaProducto.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>(),
                                                              listaDescuentosSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto) && !x.Destino.Equals(string.Empty)).Distinct().ToList<DescuentoSAPBE>().ToCollection<DescuentoSAPBE>(),
                                                              listaTarifasSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TarifaSAPBE>().ToCollection<TarifaSAPBE>(),
                                                              listaGrados.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                                                              listaPenalizaciones.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PenalizacionRegularidadProductoBE>().ToCollection<PenalizacionRegularidadProductoBE>(),
                                                              listaRegularidades.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                                                              listaUmbrales.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                                                              listaRangosPoblacionD2.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<RangoPoblacionD2BE>().ToCollection<RangoPoblacionD2BE>(),
                                                              ref listaVA,
                                                              listaProductoInternacional.FirstOrDefault(x => x.Producto.Equals(itemProductoBE.CodProducto) && x.Anexo.Equals(itemProductoBE.CodAnexoSAP) && x.VA.Trim() == ""),
                                                              listaPrecioPor.Where(x => x.ProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PrecioPorBE>().ToCollection<PrecioPorBE>(),
                                                              listaRelacionProductos,
                                                              listaTipologiasCliente.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TipologiaClienteBE>().ToCollection<TipologiaClienteBE>()
                                                                );


                    }
                    else if (itemProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Publicorreo)))
                    {
                        //JCNS. DTO-MAX. incluyo listaTipologiasCliente
                        //this.GuardarDefinicionProductoPublicorreo(nuow,
                        //                                      itemProductoBE,
                        //                                      listaEsquemaProducto.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>(),
                        //                                      listaDescuentosSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto) ).Distinct().ToList<DescuentoSAPBE>().ToCollection<DescuentoSAPBE>(),
                        //                                      listaTarifasSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TarifaSAPBE>().ToCollection<TarifaSAPBE>(),
                        //                                      listaPuntos.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                        //                                      listaGrados.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                        //                                      listaPenalizaciones.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PenalizacionRegularidadProductoBE>().ToCollection<PenalizacionRegularidadProductoBE>(),
                        //                                      listaRegularidades.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                        //                                      ref listaVA,
                        //                                      listaProductoInternacional.FirstOrDefault(x => x.Producto.Equals(itemProductoBE.CodProducto) && x.Anexo.Equals(itemProductoBE.CodAnexoSAP) && x.VA.Trim() == ""),
                        //                                      listaPrecioPor.Where(x => x.ProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PrecioPorBE>().ToCollection<PrecioPorBE>(),

                        //                                      listaRelacionProductos,
                        //                                      listaCaracteristicas.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<ConfiguracionCaracteristicaBE>().ToCollection<ConfiguracionCaracteristicaBE>());
                        this.GuardarDefinicionProductoPublicorreo(nuow,
                                                              itemProductoBE,
                                                              listaEsquemaProducto.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>(),
                                                              listaDescuentosSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto)).Distinct().ToList<DescuentoSAPBE>().ToCollection<DescuentoSAPBE>(),
                                                              listaTarifasSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TarifaSAPBE>().ToCollection<TarifaSAPBE>(),
                                                              listaPuntos.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                                                              listaGrados.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                                                              listaPenalizaciones.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PenalizacionRegularidadProductoBE>().ToCollection<PenalizacionRegularidadProductoBE>(),
                                                              listaRegularidades.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                                                              ref listaVA,
                                                              listaProductoInternacional.FirstOrDefault(x => x.Producto.Equals(itemProductoBE.CodProducto) && x.Anexo.Equals(itemProductoBE.CodAnexoSAP) && x.VA.Trim() == ""),
                                                              listaPrecioPor.Where(x => x.ProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PrecioPorBE>().ToCollection<PrecioPorBE>(),

                                                              listaRelacionProductos,
                                                              listaCaracteristicas.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<ConfiguracionCaracteristicaBE>().ToCollection<ConfiguracionCaracteristicaBE>(),
                                                              listaTipologiasCliente.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TipologiaClienteBE>().ToCollection<TipologiaClienteBE>()
                                                              );
                    }
                    else if (itemProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Etiquetas)))
                    {
                        //JCNS. DTO-MAX. incluyo listaTipologiasCliente
                        //this.GuardarDefinicionProductoEtiquetas(nuow,
                        //                                    itemProductoBE,
                        //                                    listaEsquemaProducto.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>(),
                        //                                    listaDescuentosSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto) && !x.Destino.Equals(string.Empty)).Distinct().ToList<DescuentoSAPBE>().ToCollection<DescuentoSAPBE>(),
                        //                                    listaTarifasSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TarifaSAPBE>().ToCollection<TarifaSAPBE>(),
                        //                                    listaCaracteristicas.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<ConfiguracionCaracteristicaBE>().ToCollection<ConfiguracionCaracteristicaBE>(),
                        //                                    ref listaVA,
                        //                                    listaProductoInternacional.FirstOrDefault(x => x.Producto.Equals(itemProductoBE.CodProducto) && x.Anexo.Equals(itemProductoBE.CodAnexoSAP) && x.VA.Trim() == ""),
                        //                                    listaPrecioPor.Where(x => x.ProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PrecioPorBE>().ToCollection<PrecioPorBE>(),

                        //                                    listaRelacionProductos);
                        this.GuardarDefinicionProductoEtiquetas(nuow,
                                                            itemProductoBE,
                                                            listaEsquemaProducto.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>(),
                                                            listaDescuentosSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto) && !x.Destino.Equals(string.Empty)).Distinct().ToList<DescuentoSAPBE>().ToCollection<DescuentoSAPBE>(),
                                                            listaTarifasSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TarifaSAPBE>().ToCollection<TarifaSAPBE>(),
                                                            listaCaracteristicas.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<ConfiguracionCaracteristicaBE>().ToCollection<ConfiguracionCaracteristicaBE>(),
                                                            ref listaVA,
                                                            listaProductoInternacional.FirstOrDefault(x => x.Producto.Equals(itemProductoBE.CodProducto) && x.Anexo.Equals(itemProductoBE.CodAnexoSAP) && x.VA.Trim() == ""),
                                                            listaPrecioPor.Where(x => x.ProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PrecioPorBE>().ToCollection<PrecioPorBE>(),

                                                            listaRelacionProductos,
                                                              listaTipologiasCliente.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TipologiaClienteBE>().ToCollection<TipologiaClienteBE>()
                                                            );
                    }
                    else if (itemProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Libros)))
                    {
                        //JCNS. DTO-MAX. incluyo listaTipologiasCliente
                        //this.GuardarDefinicionProductoLibros(nuow,
                        //                                 itemProductoBE,
                        //                                 listaEsquemaProducto.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>(),
                        //                                 listaDescuentosSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto) && !x.Destino.Equals(string.Empty)).Distinct().ToList<DescuentoSAPBE>().ToCollection<DescuentoSAPBE>(),
                        //                                 listaTarifasSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TarifaSAPBE>().ToCollection<TarifaSAPBE>(),
                        //                                 listaPuntos.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                        //                                 listaGrados.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                        //                                 listaPenalizaciones.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PenalizacionRegularidadProductoBE>().ToCollection<PenalizacionRegularidadProductoBE>(),
                        //                                 listaRegularidades.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                        //                                 ref listaVA,
                        //                                 listaProductoInternacional.FirstOrDefault(x => x.Producto.Equals(itemProductoBE.CodProducto) && x.Anexo.Equals(itemProductoBE.CodAnexoSAP) && x.VA.Trim() == ""),
                        //                                 listaPrecioPor.Where(x => x.ProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PrecioPorBE>().ToCollection<PrecioPorBE>(),


                        //                                 listaRelacionProductos,
                        //                                 listaCaracteristicas.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<ConfiguracionCaracteristicaBE>().ToCollection<ConfiguracionCaracteristicaBE>());
                        this.GuardarDefinicionProductoLibros(nuow,
                                                         itemProductoBE,
                                                         listaEsquemaProducto.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>(),
                                                         listaDescuentosSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto) && !x.Destino.Equals(string.Empty)).Distinct().ToList<DescuentoSAPBE>().ToCollection<DescuentoSAPBE>(),
                                                         listaTarifasSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TarifaSAPBE>().ToCollection<TarifaSAPBE>(),
                                                         listaPuntos.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                                                         listaGrados.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                                                         listaPenalizaciones.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PenalizacionRegularidadProductoBE>().ToCollection<PenalizacionRegularidadProductoBE>(),
                                                         listaRegularidades.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                                                         ref listaVA,
                                                         listaProductoInternacional.FirstOrDefault(x => x.Producto.Equals(itemProductoBE.CodProducto) && x.Anexo.Equals(itemProductoBE.CodAnexoSAP) && x.VA.Trim() == ""),
                                                         listaPrecioPor.Where(x => x.ProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PrecioPorBE>().ToCollection<PrecioPorBE>(),


                                                         listaRelacionProductos,
                                                         listaCaracteristicas.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<ConfiguracionCaracteristicaBE>().ToCollection<ConfiguracionCaracteristicaBE>(),
                                                              listaTipologiasCliente.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TipologiaClienteBE>().ToCollection<TipologiaClienteBE>()
                                                         );
                    }
                    else if (itemProductoBE.ModeloDescuento.Equals(ModeloDescuentoEnum.ObtenerNombreEnum(ModeloDescuento.Publicaciones)))
                    {
                        //JCNS. DTO-MAX. incluyo listaTipologiasCliente
                        //this.GuardarDefinicionProductoPublicaciones(nuow,
                        //                                        itemProductoBE,
                        //                                        listaEsquemaProducto.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>(),
                        //                                        listaDescuentosSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto) && !x.Destino.Equals(string.Empty)).Distinct().ToList<DescuentoSAPBE>().ToCollection<DescuentoSAPBE>(),
                        //                                        listaTarifasSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TarifaSAPBE>().ToCollection<TarifaSAPBE>(),
                        //                                        listaPuntos.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                        //                                        listaGrados.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                        //                                        listaPenalizaciones.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PenalizacionRegularidadProductoBE>().ToCollection<PenalizacionRegularidadProductoBE>(),
                        //                                        listaRegularidades.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                        //                                        ref listaVA,
                        //                                        listaProductoInternacional.FirstOrDefault(x => x.Producto.Equals(itemProductoBE.CodProducto) && x.Anexo.Equals(itemProductoBE.CodAnexoSAP) && x.VA.Trim() == ""),
                        //                                        listaPrecioPor.Where(x => x.ProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PrecioPorBE>().ToCollection<PrecioPorBE>(),
                        //                                        listaRelacionProductos,
                        //                                        listaCaracteristicas.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<ConfiguracionCaracteristicaBE>().ToCollection<ConfiguracionCaracteristicaBE>());
                        this.GuardarDefinicionProductoPublicaciones(nuow,
                                                                itemProductoBE,
                                                                listaEsquemaProducto.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>(),
                                                                listaDescuentosSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto) && !x.Destino.Equals(string.Empty)).Distinct().ToList<DescuentoSAPBE>().ToCollection<DescuentoSAPBE>(),
                                                                listaTarifasSAP.Where(x => x.Producto.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TarifaSAPBE>().ToCollection<TarifaSAPBE>(),
                                                                listaPuntos.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                                                                listaGrados.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                                                                listaPenalizaciones.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PenalizacionRegularidadProductoBE>().ToCollection<PenalizacionRegularidadProductoBE>(),
                                                                listaRegularidades.FirstOrDefault(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)),
                                                                ref listaVA,
                                                                listaProductoInternacional.FirstOrDefault(x => x.Producto.Equals(itemProductoBE.CodProducto) && x.Anexo.Equals(itemProductoBE.CodAnexoSAP) && x.VA.Trim() == ""),
                                                                listaPrecioPor.Where(x => x.ProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<PrecioPorBE>().ToCollection<PrecioPorBE>(),
                                                                listaRelacionProductos,
                                                                listaCaracteristicas.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<ConfiguracionCaracteristicaBE>().ToCollection<ConfiguracionCaracteristicaBE>(),
                                                                listaTipologiasCliente.Where(x => x.CodProductoSAP.Equals(itemProductoBE.CodProducto)).Distinct().ToList<TipologiaClienteBE>().ToCollection<TipologiaClienteBE>()
                                                                );
                    }


                    this.EliminarCoeficientesProducto(itemProductoBE.idProducto, usuario, nuow);

                    //Se eliminan las potencialidades del producto antiguo y de los valores añadidos del producto antiguo
                    this.EliminarPotencialidadesProducto(itemProductoBE.idProducto, usuario, nuow);

                    //Se salva el contexto global de la descarga
                    nuow.Save();
                    //JCNS. TARIFAS 2020
                    RegistrarAccionesSimulador.GuardarTraza("FIN    Actualizar Producto " + itemProductoBE.CodAnexoSAP + " - " + itemProductoBE.CodProducto);
                }
            }

        }






        #endregion

        #region Métodos Obtener genéricos

        /// <summary>
        /// Función que devuelve los tramos rellenos con valor 
        /// </summary>
        /// <param name="idProductoOferta"></param>
        /// <param name="idDestino"></param>
        /// <returns></returns>
        public Collection<ConfiguracionValorAnadidoBE> ObtenerVARellenos(Guid idProductoOferta)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ConfiguracionValorAnadidoPersistence persistence = new ConfiguracionValorAnadidoPersistence(uow);
                return persistence.ObtenerListaConfiguracionVA(idProductoOferta);
            }
        }

        /// <summary>
        /// Método que realiza una copia de los datos de configuracion de los valores añadidos de un producto oferta
        /// </summary>
        /// <param name="idProductoOfertaOrigen">Identificador del producto oferta origen</param>
        /// <param name="idProductoOfertaDestino">Identificador del producto oferta destino</param>
        /// <returns>Colección de entidades ConfiguracionValorAnadidoBE</returns>
        public Collection<ConfiguracionValorAnadidoBE> ObtenerCopiaListaConfiguracionVA(Guid idProductoOfertaOrigen, Guid idProductoOfertaDestino)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ConfiguracionValorAnadidoPersistence persistence = new ConfiguracionValorAnadidoPersistence(uow);
                return persistence.ObtenerCopiaListaConfiguracionVA(idProductoOfertaOrigen, idProductoOfertaDestino);
            }
        }

        /// <summary>
        /// Método que realiza una copia de los datos de configuracion de los valores añadidos de un producto oferta
        /// </summary>
        /// <param name="idProductoOfertaOrigen">Identificador del producto oferta origen</param>
        /// <param name="idProductoOfertaDestino">Identificador del producto oferta destino</param>
        /// <returns>Colección de entidades ConfiguracionValorAnadidoBE</returns>
        public Collection<ConfiguracionValorAnadidoBE> ObtenerCopiaListaConfiguracionVAActualizacionProducto(Guid idProductoOfertaOrigen, Guid idProductoOfertaDestino)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ConfiguracionValorAnadidoPersistence persistence = new ConfiguracionValorAnadidoPersistence(uow);
                return persistence.ObtenerCopiaListaConfiguracionVAActualizacionProducto(idProductoOfertaOrigen, idProductoOfertaDestino);
            }
        }

        /// <summary>
        /// Método que obtiene la lista de registros de la tabla ConfiguracionDestinoOferta correspondiente con el producto oferta pasado por parámetro
        /// </summary>
        /// <param name="idProductoOferta">Identificador del producto oferta</param>
        /// <returns>Lista de ConfiguracionDestinoOfertaBE</returns>
        public Collection<ConfiguracionDestinoOfertaBE> ObtenerConfiguracionDestinoOferta(Guid idProductoOferta)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                DestinoPersistence persistence = new DestinoPersistence(uow);
                return persistence.ObtenerConfiguracionDestinoOferta(idProductoOferta);
            }
        }

        /// <summary>
        /// Método que obtiene todos los registros de ConfiguraciónDestinoOferta del listado de productosOferta pasado por parámetro
        /// y cuyas distribuciones sean mayores que 0
        /// </summary>
        /// <param name="listaIdsProductoOferta">Listado de identificadores de productosOferta</param>
        /// <returns>Colección de entidades ConfiguracionDestinoOfertaBE</returns>
        public Collection<ConfiguracionDestinoOfertaBE> ObtenerConfiguracionDestinoOferta(Collection<Guid> listaIdsProductoOferta)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                DestinoPersistence persistence = new DestinoPersistence(uow);
                return persistence.ObtenerConfiguracionDestinoOferta(listaIdsProductoOferta);
            }
        }

        /// <summary>
        /// Método que copia la configuración de destinos de un producto oferta
        /// </summary>
        /// <param name="idProductoOriginal">Identificador del producto oferta original</param>
        /// <param name="idProductoDestino">Identificador del producto oferta destino</param>
        /// <returns>Colección de entidades ConfiguracionListaPreciosBE</returns>
        public Collection<ConfiguracionDestinoOfertaBE> CopiarConfiguracionDestinoProductoOferta(Guid idProductoOfertaOriginal, Guid idProductoOfertaDestino)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                DestinoPersistence persistence = new DestinoPersistence(uow);
                return persistence.CopiarConfiguracionDestinoProductoOferta(idProductoOfertaOriginal, idProductoOfertaDestino);
            }
        }

        /// <summary>
        /// Método que copia la configuración de destinos que aparece en una nueva definicion de un producto oferta
        /// </summary>
        /// <param name="idProductoOriginal">Identificador del producto oferta original</param>
        /// <param name="idProductoDestino">Identificador del producto oferta destino</param>
        /// <returns>Colección de entidades ConfiguracionListaPreciosBE</returns>
        public Collection<ConfiguracionDestinoOfertaBE> CopiarConfiguracionDestinoProductoOfertaActualizacionProducto(Guid idProductoOfertaOriginal, Guid idProductoOfertaDestino)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                DestinoPersistence persistence = new DestinoPersistence(uow);
                return persistence.CopiarConfiguracionDestinoProductoOfertaActualizacionProducto(idProductoOfertaOriginal, idProductoOfertaDestino);
            }
        }

        /// <summary>
        /// Método que obtiene el registro de la tabla ConfiguracionGradoOferta correspondiente con el grado y el producto
        /// oferta pasados por parámetro
        /// </summary>
        /// <param name="idDestino">Identificador del grado</param>
        /// <param name="idProductoOferta">Identificador del producto oferta</param>
        /// <returns>Entidad ConfiguracionGradoOfertaBE</returns>
        public ConfiguracionGradoOfertaBE ObtenerConfiguracionGradoOferta(Guid idGrado, Guid idProductoOferta)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ConfiguracionGradoOfertaPersistence persistence = new ConfiguracionGradoOfertaPersistence(uow);
                return persistence.ObtenerConfiguracionGradoOferta(idGrado, idProductoOferta);
            }
        }

        /// <summary>
        /// Método que obtiene el registro de la tabla ConfiguracionPuntoOferta correspondiente con el punto y el producto
        /// oferta pasados por parámetro
        /// </summary>
        /// <param name="idDestino">Identificador del grado</param>
        /// <param name="idProductoOferta">Identificador del producto oferta</param>
        /// <returns>Entidad ConfiguracionGradoOfertaBE</returns>
        public ConfiguracionPuntoOfertaBE ObtenerConfiguracionPuntoOferta(Guid idPunto, Guid idProductoOferta)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ConfiguracionPuntoOfertaPersistence persistence = new ConfiguracionPuntoOfertaPersistence(uow);
                return persistence.ObtenerConfiguracionPuntoOferta(idPunto, idProductoOferta);
            }
        }

        /// <summary>
        /// Método que obtiene todos los registros de la tabla configuracionPuntoOferta de un listado de productosOferta pasado por parámetro.
        /// </summary>
        /// <param name="listaIdsProductosOferta">Listado de identificadores de productosOferta</param>
        /// <returns>Colección de entidades ConfiguracionPuntoOfertaBE</returns>
        public Collection<ConfiguracionPuntoOfertaBE> ObtenerConfiguracionPuntoOferta(Collection<Guid> listaIdsProductosOferta)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ConfiguracionPuntoOfertaPersistence persistence = new ConfiguracionPuntoOfertaPersistence(uow);
                return persistence.ObtenerConfiguracionPuntoOferta(listaIdsProductosOferta);
            }
        }

        // <summary>
        /// Método que copia la configuración de puntos de un producto oferta
        /// </summary>
        /// <param name="idProductoOfertaOrigen">Identificador del producto oferta origen</param>
        /// <param name="idProductoOfertaDestino">Identificador del producto oferta destino</param>
        /// <returns>Colección de entidades ConfiguracionListaPreciosBE</returns>
        public Collection<ConfiguracionPuntoOfertaBE> CopiarConfiguracionPuntoProductoOferta(Guid idProductoOfertaOrigen, Guid idProductoOfertaDestino)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ConfiguracionPuntoOfertaPersistence persistence = new ConfiguracionPuntoOfertaPersistence(uow);
                return persistence.CopiarConfiguracionPuntoProductoOferta(idProductoOfertaOrigen, idProductoOfertaDestino);
            }
        }

        /// <summary>
        /// Método que obtiene la lista de registros de la tabla ConfiguracionTramoOferta correspondiente con el producto
        /// oferta pasado por parámetro
        /// </summary>
        /// <param name="idTramo">Identificador del tramo</param>
        /// <param name="idProductoOferta">Identificador del producto oferta</param>
        /// <returns>Entidad ConfiguracionTramoOfertaBE</returns>
        public Collection<ConfiguracionTramoOfertaBE> ObtenerConfiguracionTramoOferta(Guid idProductoOferta)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                TramoPersistence persistence = new TramoPersistence(uow);
                return persistence.ObtenerConfiguracionTramoOferta(idProductoOferta);
            }
        }

        /// <summary>
        /// Método que obtiene todos los registros de la tabla ConfiguraciónTramoOferta correspondientes a los productosOferta
        /// pasados por parámetro y cuya distribución sea mayor que 0
        /// </summary>
        /// <param name="listaIdsProductosOferta">Listado de identificadores de productosOferta</param>
        /// <returns>Entidad ConfiguracionTramoOfertaBE</returns>
        public Collection<ConfiguracionTramoOfertaBE> ObtenerConfiguracionTramoOferta(Collection<Guid> listaIdsProductosOferta)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                TramoPersistence persistence = new TramoPersistence(uow);
                return persistence.ObtenerConfiguracionTramoOferta(listaIdsProductosOferta);
            }
        }

        /// <summary>
        /// Método que devuelve todos los registros de la tabla ConfiguracionTramoOferta
        /// </summary>
        /// <returns>Colección de entidades ConfiguracionTramoOfertaBE</returns>
        public Collection<ConfiguracionTramoOfertaBE> ObtenerConfiguracionTramoOfertaCompleta()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                TramoPersistence persistence = new TramoPersistence(uow);
                return persistence.ObtenerConfiguracionTramoOfertaCompleta();
            }
        }

        /// <summary>
        /// Método que devuelve todos los registros de la tabla ConfiguracionTramoOferta correspondientes a los productos de la lista
        /// </summary>
        /// <returns>Colección de entidades ConfiguracionTramoOfertaBE</returns>
        public Collection<ConfiguracionTramoOfertaBE> ObtenerConfiguracionTramoOferta(Collection<Guid> listaIdsProductosOferta, IUnitOfWork uow)
        {
            TramoPersistence persistence = new TramoPersistence(uow);
            return persistence.ObtenerConfiguracionTramoOferta(listaIdsProductosOferta, uow);
        }

        /// <summary>
        /// Método que copia la configuración de tramos de un producto oferta
        /// </summary>
        /// <param name="idProductoOfertaOrigen">Identificador del productoOferta origen</param>
        /// <param name="idProductoOfertaDestino">Identificador del producto Oferta destino</param>
        /// <returns>Entidad ConfiguracionTramoOfertaBE</returns>
        public Collection<ConfiguracionTramoOfertaBE> CopiarConfiguracionTramoProductoOferta(Guid idProductoOfertaOrigen, Guid idProductoOfertaDestino)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                TramoPersistence persistence = new TramoPersistence(uow);
                return persistence.CopiarConfiguracionTramoProductoOferta(idProductoOfertaOrigen, idProductoOfertaDestino);
            }
        }

        /// <summary>
        /// Método que copia la configuración de tramos para la actualizacion de un producto oferta
        /// </summary>
        /// <param name="idProductoOfertaOrigen">Identificador del productoOferta origen</param>
        /// <param name="idProductoOfertaDestino">Identificador del producto Oferta destino</param>
        /// <param name="destinosOriginal">Lista de destinos del producto oferta origen</param>
        /// <param name="destinosOriginal">Lista de destinos del producto Oferta destino</param>
        /// <returns>Entidad ConfiguracionTramoOfertaBE</returns>
        public Collection<ConfiguracionTramoOfertaBE> CopiarConfiguracionTramoProductoOfertaActualizacionProducto(Guid idProductoOfertaOrigen, Guid idProductoOfertaDestino, Collection<DestinoBE> destinosOriginal, Collection<DestinoBE> destinosCopia)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                TramoPersistence persistence = new TramoPersistence(uow);
                return persistence.CopiarConfiguracionTramoProductoOfertaActualizacionProducto(idProductoOfertaOrigen, idProductoOfertaDestino, destinosOriginal, destinosCopia);
            }
        }

        /// <summary>
        /// Método que obtiene el registro de la tabla ConfiguracionValorAnadido correspondiente con el valor añadido y el 
        /// producto oferta pasados por parámetro
        /// </summary>
        /// <param name="idValorAnadidoProducto">Identificador del valor añadido producto</param>
        /// <param name="idProductoOferta">Identificador del producto oferta</param>
        /// <returns>Entidad ConfiguracionValorAnadidoBE</returns>
        public ConfiguracionValorAnadidoBE ObtenerConfiguracionValorAnadido(Guid idValorAnadidoProducto, Guid idProductoOferta)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ValorAnadidoProductoPersistence persistence = new ValorAnadidoProductoPersistence(uow);
                return persistence.ObtenerConfiguracionValorAnadido(idValorAnadidoProducto, idProductoOferta);
            }
        }

        /// <summary>
        /// Método que obtiene los registros de la tabla ConfiguracionValorAnadido correspondientes con los valores añadidos y el 
        /// producto oferta pasados por parámetro
        /// </summary>
        /// <param name="idValorAnadidoProducto">Liste de identificadores del valor añadido producto</param>
        /// <param name="idProductoOferta">Identificador del producto oferta</param>
        /// <returns>Colección de entidades ConfiguracionValorAnadidoBE</returns>
        public Collection<ConfiguracionValorAnadidoBE> ObtenerConfiguracionesValorAnadido(Collection<Guid> idsValorAnadido, Guid idProductoOferta)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ValorAnadidoProductoPersistence persistence = new ValorAnadidoProductoPersistence(uow);
                return persistence.ObtenerConfiguracionesValorAnadido(idsValorAnadido, idProductoOferta);
            }
        }

        #endregion

        #region Méodos Eliminar

        /// <summary>
        /// Elimina las definiciones de los productos que están obsoletos
        /// </summary>
        public void EliminarDefinicionesObsoletas()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProductoPersistence persistence = new ProductoPersistence(uow);

                persistence.MarcarAnexo90Obsoleto();
                uow.Save();

                persistence.EliminarDefinicionesObsoletas();

                persistence.EliminarAnexo90();
                uow.Save();
            }
        }

        #endregion

        #endregion

        #region Métodos Privados

        /// <summary>
        /// Método que marca establece la fecha de validez hasta para los productos pasados por parámetro, marcando su definición como obsoleta
        /// </summary>
        /// <param name="listaProductos">Lista de ProductoBE</param>
        /// <param name="uow">Objeto transaccional</param>
        private void MarcarDefinicionProductoObsoleta(ProductoBE itemProductosBE, IUnitOfWork uow)
        {
            ProductoPersistence productoPersistence = new ProductoPersistence(uow);
            ProductoBE productoDB = productoPersistence.ObtenerProductoByIdAnexoProducto(itemProductosBE.idAnexoProducto);

            // El producto no puede ser nulo, no obstante se comprueba
            if (productoDB != null)
            {
                //Se guarda la fecha de validez final
                productoDB.ValidezHasta = DateTime.Now;
                productoDB.ActualizacionPendiente = false;
                productoPersistence.InsertUpdateProducto(productoDB);
            }

        }


        /// <summary>
        /// guarda la definicion de un producto de etiquetas
        /// </summary>
        /// <param name="uow">contexto de la base de datos</param>
        /// <param name="itemProductoAnterior">producto del que se quiere actualizar la definición</param>
        /// <param name="esqueletoProducto">esquema de los prodictos y VA que tiene el proyecto</param>
        /// <param name="listaDescuento">lista de descuentos que tiene el producto</param>
        /// <param name="listaTarifas">lista de tarifas asociadas al producto y los VA</param>
        /// <param name="listaCaracteristicas">lista de características con los valores añadidos que tiene el producto</param>
        /// <param name="listaVA">lista de los VA que actualmente están almacenados en BBDD</param>
        private void GuardarDefinicionProductoEtiquetas(IUnitOfWork uow, ProductoBE itemProductoAnterior, Collection<EsquemaProductoBE> esqueletoProducto, Collection<DescuentoSAPBE> listaDescuento, Collection<TarifaSAPBE> listaTarifas, Collection<ConfiguracionCaracteristicaBE> listaCaracteristicas, ref Collection<ValorAnadidoBE> listaVA, InternacionalBE itemInternacional, Collection<PrecioPorBE> listaPrecioPor, Collection<RelacionProductosBE> listaRelacionProductos, Collection<TipologiaClienteBE> listaTipologias)
        {
            #region variables usadas para el guardado

            ProductoPersistence productoPersistence = new ProductoPersistence(uow);
            CaracteristicaPersistence caracteristicaPersistence = new CaracteristicaPersistence(uow);
            TarifaPersistence tarifaPersistence = new TarifaPersistence(uow);
            DescuentoPersistence descuentoPersistence = new DescuentoPersistence(uow);
            ValorAnadidoPersistence vaPersistence = new ValorAnadidoPersistence(uow);
            ValorAnadidoProductoPersistence vapPersistence = new ValorAnadidoProductoPersistence(uow);
            TipoClientePersistence tipoClientePersistence = new TipoClientePersistence(uow);
            Collection<TipoClienteBE> tiposClientesDB = tipoClientePersistence.ObtenerTiposClientes();
            RelacionProductosPersistence relacionPersistence = new RelacionProductosPersistence(uow);


            //JCNS. DTO-MAX. incluyo listaTipologiasCliente
            TipologiaClientePersistence tipologiaPersistence = new TipologiaClientePersistence(uow);

            #endregion

            #region proceso de guardado

            #region Producto

            //creo el nuevo registro en la tabla producto            
            Guid idNuevoProducto = Guid.NewGuid();

            //Se guarda como un registro nuevo.                
            productoPersistence.InsertProducto(new ProductoBE
            {
                idProducto = idNuevoProducto,
                idAnexoProducto = itemProductoAnterior.idAnexoProducto,
                ValidezDesde = DateTime.Now,
                Regularidad = null,
                ActualizacionPendiente = false,
                UmbralD2 = null,
                CodProducto = itemProductoAnterior.CodProducto,
                CodAnexoSAP = itemProductoAnterior.CodAnexoSAP,
                Internacional = itemInternacional != null ? true : false,
            });

            #endregion

            #region Tipologia Producto
            //JCNS. DTO-MAX. incluyo listaTipologiasCliente. NO LO HACÍA EN ETIQUETAS

            //guardo la tipología del producto nuevo
            foreach (TipologiaClienteBE itemTipologia in listaTipologias.Where(x => string.IsNullOrWhiteSpace(x.CodValorAnadidoSAP)))
            {
                TipoClienteBE tipoClienteDB = tiposClientesDB.FirstOrDefault(x => x.CodTipoCliente.Equals(itemTipologia.TipoCliente));
                if (tipoClienteDB != null)
                {
                    tipologiaPersistence.InsertTipologiaProducto(new TipologiaClienteBE()
                    {
                        idTipologiaProducto = Guid.NewGuid(),
                        idProducto = idNuevoProducto,
                        idTipoCliente = tipoClienteDB.idTipoCliente,
                        Potencialidad = itemTipologia.Potencialidad,
                        Facturacion = itemTipologia.Facturacion
                    });
                }
            }

            #endregion

            #region Características Producto

            foreach (ConfiguracionCaracteristicaBE caracteristica in listaCaracteristicas)
            {
                bool existeCaracteristica = false;

                //Obtengo la característica asociada
                caracteristica.idCaracteristica = caracteristicaPersistence.ObtenerIDCaracteristica(caracteristica.NombreCaracteristica, caracteristica.DescripcionCaracteristica, out existeCaracteristica);
                caracteristica.idProducto = idNuevoProducto;

                if (!existeCaracteristica)
                {
                    //Sólo inserto los valores de la característica si no existía previamente
                    foreach (var item in caracteristica.ListaValores)
                    {
                        caracteristicaPersistence.InsertUpdateValor(item.Valor, item.Descripcion, caracteristica.idCaracteristica);
                    }
                }

                //cuando ya tengo todo el proceso anterior, añado un nuevo registro en bbdd
                caracteristicaPersistence.InsertCaracteristicaProducto(caracteristica);
            }

            #endregion Características Producto

            #region Tarifa y descuentos

            //esqueleto producto
            Collection<EsquemaProductoBE> auxDestinos = esqueletoProducto.Where(x => !x.CodDestinoSAP.Equals("S/D") && string.IsNullOrWhiteSpace(x.CodTramoSAP) && string.IsNullOrWhiteSpace(x.VASAP)).ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>();
            if (auxDestinos.Count.Equals(0))
            {
                //es un producto sin destinos
                #region Tarifas

                foreach (TarifaSAPBE objTarifa in listaTarifas.Where(x => string.IsNullOrWhiteSpace(x.VA)))
                {
                    //Se guarda la tarifa
                    TarifaBE tarifaProductoDB = new TarifaBE();

                    tarifaProductoDB.idTarifa = Guid.NewGuid();
                    tarifaProductoDB.idProducto = idNuevoProducto;
                    tarifaProductoDB.TipoPrecio = objTarifa.TipoPrecio;

                    decimal auxTarifa = 0;
                    decimal.TryParse(objTarifa.Tarifa.ToString(), out auxTarifa);

                    tarifaProductoDB.Tarifa = auxTarifa;
                    tarifaProductoDB.Moneda = objTarifa.Moneda;

                    tarifaProductoDB.Combinacion = objTarifa.Combinacion;

                    tarifaPersistence.InsertTarifaProductoSinDestino(tarifaProductoDB);
                }


                #endregion

                #region Descuentos

                DescuentoSAPBE itemDescuento = listaDescuento.FirstOrDefault(x => string.IsNullOrWhiteSpace(x.VA));
                if (itemDescuento != null)
                {
                    TipoClienteBE tipoClienteDB = tiposClientesDB.FirstOrDefault(x => x.CodTipoCliente.Equals("A"));
                    if (tipoClienteDB != null)
                    {
                        descuentoPersistence.InsertDescuentoProductoSinDestino(new DescuentoBE()
                        {
                            idDescuento = Guid.NewGuid(),
                            idProducto = idNuevoProducto,
                            idTipoCliente = tipoClienteDB.idTipoCliente,
                            DtoMax = itemDescuento.DtoMax,
                            DtoMaxTDC = itemDescuento.DtoMaxTDC
                        });
                    }

                }

                #endregion
            }

            #endregion Tarifa y descuentos

            #region VA

            Collection<EsquemaProductoBE> auxVAs = esqueletoProducto.Where(x => string.IsNullOrWhiteSpace(x.CodTramoSAP) && string.IsNullOrWhiteSpace(x.CodDestinoSAP) && !string.IsNullOrWhiteSpace(x.VASAP)).ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>();
            foreach (EsquemaProductoBE itemVA in auxVAs)
            {
                //solo se tratan los que realmente son VA...
                if ((!itemVA.VASAP.Equals("S0098")) && (!itemVA.VASAP.Equals("S0099")) && (!itemVA.VASAP.Equals("S0134")))
                {
                    #region Añadir VA al maestro de VA

                    ValorAnadidoBE valorAnadidoDB = listaVA.FirstOrDefault(x => x.CodValorAnadidoSAP.Equals(itemVA.VASAP));
                    if (valorAnadidoDB == null || String.IsNullOrEmpty(valorAnadidoDB.Descripcion))
                    {
                        Boolean isUpdate = valorAnadidoDB != null && String.IsNullOrEmpty(valorAnadidoDB.Descripcion);

                        //No está en base de datos, hay que añadirlo e insertarlo
                        if (valorAnadidoDB == null)
                        {
                            valorAnadidoDB = new ValorAnadidoBE()
                            {
                                idValorAnadido = Guid.NewGuid(),
                                CodValorAnadidoSAP = itemVA.VASAP,
                                Descripcion = itemVA.DescripcionVA
                            };
                        }
                        else
                        {
                            //Existe pero le faltan valores
                            valorAnadidoDB.Descripcion = itemVA.DescripcionVA;
                        }

                        // Lo rellenamos con FirstOrDefault porque en los preciospor del VA dado tanto
                        // negociable por precio cierto (true, false) como modalidad de negociacion (individual, general)
                        // tienen que ser las mismas para todos
                        PrecioPorBE precioPorVA = listaPrecioPor.FirstOrDefault(x => x.ValorAnadidoSAP.Equals(itemVA.VASAP));
                        if (precioPorVA != null)
                        {
                            valorAnadidoDB.EsParametrizable = true;
                            valorAnadidoDB.NegociableAPC = precioPorVA.NegociablePorPrecioCierto;
                            valorAnadidoDB.ModalidadNegociacion = precioPorVA.ModalidadNegociacionTarifa;
                        }

                        if (isUpdate)
                        {
                            vaPersistence.InsertUpdateValorAnadido(valorAnadidoDB);
                        }
                        else
                        {
                            vaPersistence.InsertValorAnadido(valorAnadidoDB);
                            listaVA.Add(valorAnadidoDB);

                        }
                    }

                    #endregion

                    #region Relación del VA con el producto

                    //inserto la relación del producto y del VA                    
                    ValorAnadidoAnexoProductoBE auxValorAnadidoProducto = new ValorAnadidoAnexoProductoBE()
                    {
                        idValorAnadidoAnexoProducto = Guid.NewGuid(),
                        idProducto = idNuevoProducto,
                        idValorAnadido = valorAnadidoDB.idValorAnadido
                    };
                    vapPersistence.InsertValorAnadidoProducto(auxValorAnadidoProducto);

                    #region Tarifa del VA

                    //tarifa VA                    
                    foreach (TarifaSAPBE obTarifa in listaTarifas.Where(x => x.VA.Equals(itemVA.VASAP)))
                    {
                        TarifaBE tarifaProductoDB = new TarifaBE();

                        tarifaProductoDB.idTarifa = Guid.NewGuid();
                        tarifaProductoDB.TipoPrecio = obTarifa.TipoPrecio;

                        decimal auxTarifa = 0;
                        decimal.TryParse(obTarifa.Tarifa.ToString(), out auxTarifa);

                        tarifaProductoDB.Tarifa = auxTarifa;
                        tarifaProductoDB.Moneda = obTarifa.Moneda;
                        tarifaProductoDB.idValorAnadidoProducto = auxValorAnadidoProducto.idValorAnadidoAnexoProducto;

                        tarifaPersistence.InsertTarifaVA(tarifaProductoDB);
                    }

                    #endregion

                    #endregion
                }
            }

            #endregion

            #region Relación Productos

            //Borramos las relaciones asociadas al producto
            relacionPersistence.DeleteRelacionesProdSAPSobrantes(itemProductoAnterior.CodProducto, listaRelacionProductos);

            //Insertamos las relaciones asociadas al producto
            var listaRelacionesProducto = listaRelacionProductos.Where(t => t.CodProductoSAP_A.Equals(itemProductoAnterior.CodProducto) || t.CodProductoSAP_B.Equals(itemProductoAnterior.CodProducto));

            foreach (var item in listaRelacionesProducto)
            {
                relacionPersistence.InsertRelacionProductos(item);
            }

            #endregion

            #endregion

        }


        /// <summary>
        /// Guarda la definición de un producto del modelo de Tramos
        /// </summary>
        /// <param name="uow">Contexto de la base de datos</param>
        /// <param name="itemProductoAnterior">producto del que se quieres actualizar la definicion</param>
        /// <param name="esqueletoProducto">esqueleto de los destinos y tramos del producto descargado</param>
        /// <param name="listaDescuento">lista de descuentos del producto</param>
        /// <param name="listaTarifas">lista de las tarifas del producto</param>
        /// <param name="listaTipologias">lista de las tipologias aplicables al producto</param>
        /// <param name="listaVA">lista de VA soportados por el sistema</param>
        private void GuardarDefinicionProductoTramos(IUnitOfWork uow, ProductoBE itemProductoAnterior, Collection<EsquemaProductoBE> esqueletoProducto, Collection<DescuentoSAPBE> listaDescuento, Collection<TarifaSAPBE> listaTarifas, Collection<TipologiaClienteBE> listaTipologias, Collection<ConfiguracionCaracteristicaBE> listaCaracteristicas, ref Collection<ValorAnadidoBE> listaVA, Collection<PrecioPorBE> listaPrecioPor, InternacionalBE itemInternacional, Collection<RelacionProductosBE> listaRelacionProductos)
        {
            #region variables usadas para el guardado

            ProductoPersistence productoPersistence = new ProductoPersistence(uow);
            TipologiaClientePersistence tipologiaPersistence = new TipologiaClientePersistence(uow);
            TarifaPersistence tarifaPersistence = new TarifaPersistence(uow);
            DescuentoPersistence descuentoPersistence = new DescuentoPersistence(uow);
            DestinoPersistence destinoPersistence = new DestinoPersistence(uow);
            TramoPersistence tramoPersistence = new TramoPersistence(uow);
            ValorAnadidoPersistence vaPersistence = new ValorAnadidoPersistence(uow);
            ValorAnadidoProductoPersistence vapPersistence = new ValorAnadidoProductoPersistence(uow);
            TipoClientePersistence tipoClientePersistence = new TipoClientePersistence(uow);
            CaracteristicaPersistence caracteristicaPersistence = new CaracteristicaPersistence(uow);
            Collection<TipoClienteBE> tiposClientesDB = tipoClientePersistence.ObtenerTiposClientes();
            RelacionProductosPersistence relacionPersistence = new RelacionProductosPersistence(uow);
            AgrupacionPersistence agrupacionPersistence = new AgrupacionPersistence(uow);

            #endregion

            

            #region Proceso Guardar

            #region Tratamiento de las agrupaciones de los descuentos

            foreach (var itemDestino in esqueletoProducto.Where(x => !x.CodDestinoSAP.Equals("S/D") && string.IsNullOrWhiteSpace(x.CodTramoSAP) && string.IsNullOrWhiteSpace(x.VASAP)))
            {
                foreach (TipoClienteBE itemTipoCliente in tiposClientesDB)
                {
                    int indice = 0;
                    foreach (var itemDescuento in listaDescuento.Where(x => x.Destino.Equals(itemDestino.CodDestinoSAP) && x.TipoCliente.Equals(itemTipoCliente.CodTipoCliente)).OrderBy(x => x.CodTramoDesdeDecimal))
                    {
                        itemDescuento.AgrupacionTramo = indice;
                        indice++;
                    }
                }
            }


            #endregion

            #region Producto

            //creo el nuevo registro en la tabla producto            
            Guid idNuevoProducto = Guid.NewGuid();

            //Se guarda como un registro nuevo.                
            productoPersistence.InsertProducto(new ProductoBE
            {
                idProducto = idNuevoProducto,
                idAnexoProducto = itemProductoAnterior.idAnexoProducto,
                ValidezDesde = DateTime.Now,
                Regularidad = null,
                ActualizacionPendiente = false,
                UmbralD2 = null,
                CodProducto = itemProductoAnterior.CodProducto,
                CodAnexoSAP = itemProductoAnterior.CodAnexoSAP,
                Internacional = itemInternacional != null ? true : false,
            });

            #endregion

            #region Tipologia Producto

            //guardo la tipología del producto nuevo
            foreach (TipologiaClienteBE itemTipologia in listaTipologias.Where(x => string.IsNullOrWhiteSpace(x.CodValorAnadidoSAP)))
            {
                TipoClienteBE tipoClienteDB = tiposClientesDB.FirstOrDefault(x => x.CodTipoCliente.Equals(itemTipologia.TipoCliente));
                if (tipoClienteDB != null)
                {
                    tipologiaPersistence.InsertTipologiaProducto(new TipologiaClienteBE()
                    {
                        idTipologiaProducto = Guid.NewGuid(),
                        idProducto = idNuevoProducto,
                        idTipoCliente = tipoClienteDB.idTipoCliente,
                        Potencialidad = itemTipologia.Potencialidad,
                        Facturacion = itemTipologia.Facturacion
                    });
                }
            }

            #endregion

            #region Esqueleto de Destinos y Tramos
            //recorro (en función de si tiene o no) los destinos y tramos del producto y guardo la información de las tarifas y los descuentos

            //Borramos las agrupaciones del producto, para introducir las nuevas
            Collection<AgrupacionBE> listadoAgrupaciones = new Collection<AgrupacionBE>();
            //agrupacionPersistence.BorrarAgrupacionesProducto(itemProductoAnterior.idProducto);
                 
            //esqueleto producto
            Collection<EsquemaProductoBE> auxDestinos = esqueletoProducto.Where(x => !x.CodDestinoSAP.Equals("S/D") && string.IsNullOrWhiteSpace(x.CodTramoSAP) && string.IsNullOrWhiteSpace(x.VASAP)).ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>();
            if (auxDestinos.Count.Equals(0))
            {
                //es un producto sin destinos
                #region Tarifas

                TarifaSAPBE objTarifa = listaTarifas.FirstOrDefault(x => string.IsNullOrWhiteSpace(x.VA));
                if (objTarifa != null)
                {
                    //Se guarda la tarifa
                    TarifaBE tarifaProductoDB = new TarifaBE();

                    tarifaProductoDB.idTarifa = Guid.NewGuid();
                    tarifaProductoDB.idProducto = idNuevoProducto;
                    tarifaProductoDB.TipoPrecio = objTarifa.TipoPrecio;

                    decimal auxTarifa = 0;
                    decimal.TryParse(objTarifa.Tarifa.ToString(), out auxTarifa);

                    tarifaProductoDB.Tarifa = auxTarifa;
                    tarifaProductoDB.Moneda = objTarifa.Moneda;

                    tarifaPersistence.InsertTarifaProductoSinDestino(tarifaProductoDB);
                }

                #endregion

                #region Descuentos

                foreach (DescuentoSAPBE itemDescuento in listaDescuento.Where(x => string.IsNullOrWhiteSpace(x.VA)))
                {
                    TipoClienteBE tipoClienteDB = tiposClientesDB.FirstOrDefault(x => x.CodTipoCliente.Equals(itemDescuento.TipoCliente));
                    if (tipoClienteDB != null)
                    {
                        descuentoPersistence.InsertDescuentoProductoSinDestino(new DescuentoBE()
                        {
                            idDescuento = Guid.NewGuid(),
                            idProducto = idNuevoProducto,
                            idTipoCliente = tipoClienteDB.idTipoCliente,
                            DtoMax = itemDescuento.DtoMax,
                            DtoMaxTDC = itemDescuento.DtoMaxTDC
                        });
                    }

                }

                #endregion
            }
            else
            {
                //es un producto con destinos y tramos
                Collection<EsquemaProductoBE> auxTramos = esqueletoProducto.Where(x => !string.IsNullOrWhiteSpace(x.CodTramoSAP) && string.IsNullOrWhiteSpace(x.VASAP)).ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>();

                foreach (EsquemaProductoBE itemDestino in auxDestinos.OrderBy(x => x.Orden))
                {
                    Guid idNuevoDestino = Guid.NewGuid();

                    #region destino

                    //Guardamos el destino
                    DestinoBE destinoDB = new DestinoBE()
                    {
                        idDestino = idNuevoDestino,
                        CodDestinoSAP = itemDestino.CodDestinoSAP,
                        idProducto = idNuevoProducto,
                        Orden = itemDestino.Orden
                    };
                    destinoPersistence.InsertDestino(destinoDB);

                    //Guardamos las agrupaciones del destino
                    if (itemDestino.Agrupacion != null && itemDestino.Agrupacion.Length > 0)
                    {
                        foreach (var nombreAgrupacion in itemDestino.Agrupacion)
                        {
                            var nuevaAgrupacion = listadoAgrupaciones.FirstOrDefault(t => t.Nombre.Equals(nombreAgrupacion.Replace("DEFAULT_", String.Empty)));
                            Boolean anyadirAgrupacion = false;

                            //Si no existe en la lista, la añadimos
                            if (nuevaAgrupacion == null)
                            {
                                nuevaAgrupacion = new AgrupacionBE()
                                {
                                    idAgrupacion = Guid.NewGuid(),
                                    Nombre = nombreAgrupacion,
                                    AgrupacionesDestino = new List<AgrupacionDestinoBE>()
                                };

                                anyadirAgrupacion = true;                               
                            }

                            //Indicamos si es la agrupación por defecto
                            if (nuevaAgrupacion.Nombre.Contains("DEFAULT_") || nuevaAgrupacion.AgrupacionDefecto)
                            {
                                nuevaAgrupacion.AgrupacionDefecto = true;
                                nuevaAgrupacion.Nombre = nuevaAgrupacion.Nombre.Replace("DEFAULT_", String.Empty);
                            }
                            else
                            {
                                nuevaAgrupacion.AgrupacionDefecto = false;
                            }

                            if (anyadirAgrupacion)
                                listadoAgrupaciones.Add(nuevaAgrupacion);

                            nuevaAgrupacion.AgrupacionesDestino.Add(new AgrupacionDestinoBE()
                            {
                                idAgrupacionDestino = Guid.NewGuid(),
                                idAgrupacion = nuevaAgrupacion.idAgrupacion,
                                idDestino = destinoDB.idDestino
                            });
                        }
                    }
                    #endregion

                    #region tramos sus tarifas y los descuentos de los tramos con la agrupación

                    //Collection<EsquemaProductoBE> auxTramos = esqueletoProducto.Where(x => x.CodDestinoSAP.Equals(itemDestino.CodDestinoSAP) && !string.IsNullOrWhiteSpace(x.CodTramoSAP) && string.IsNullOrWhiteSpace(x.VASAP)).ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>();
                    foreach (EsquemaProductoBE itemTramo in auxTramos.OrderBy(x => x.CodTramoSAP))
                    {
                        //guardamos el tramo
                        Guid idNuevoTramo = Guid.NewGuid();
                        TramoBE tramoDB = new TramoBE()
                        {
                            idTramo = idNuevoTramo,
                            idDestino = idNuevoDestino,
                            CodTramo = itemTramo.CodTramoSAP,
                            Descripcion = itemTramo.DescripcionTramo
                        };
                        tramoPersistence.InsertTramo(tramoDB);

                        //guardamos la tarifa del tramo
                        TarifaSAPBE obTarifa = listaTarifas.FirstOrDefault(x => x.Destino.Equals(itemDestino.CodDestinoSAP) && x.Tramo.Equals(itemTramo.CodTramoSAP));

                        //Si no tiene tarifa, comprobamos si pertenece a una Zona Padre
                        if (obTarifa == null && !String.IsNullOrEmpty(itemDestino.CodDestinoZonaSAP))
                        {
                            obTarifa = listaTarifas.FirstOrDefault(x => x.Destino.Equals(itemDestino.CodDestinoZonaSAP) && x.Tramo.Equals(itemTramo.CodTramoSAP));
                        }

                        decimal auxTarifa = 0;
                        string tipoPrecio = "ZR50(EUR)";
                        string moneda = "EUR";

                        if (obTarifa != null)
                        {
                            decimal.TryParse(obTarifa.Tarifa.ToString(), out auxTarifa);
                            tipoPrecio = obTarifa.TipoPrecio;
                            moneda = obTarifa.Moneda;
                        }

                        tarifaPersistence.InsertTarifaProducto(new TarifaBE()
                        {
                            idTarifa = Guid.NewGuid(),
                            Tarifa = auxTarifa,
                            idTramo = idNuevoTramo,
                            TipoPrecio = tipoPrecio,
                            Moneda = moneda
                        });

                        //guardamos la del descuento del tramo
                        foreach (DescuentoSAPBE itemDescuento in listaDescuento.Where(x => x.CodTramoDesdeDecimal <= itemTramo.CodTramoSAPDecimal && x.CodTramoHastaDecimal >= itemTramo.CodTramoSAPDecimal && x.Destino.Equals(itemDestino.CodDestinoSAP)))
                        {
                            TipoClienteBE tipoClienteDB = tiposClientesDB.FirstOrDefault(x => x.CodTipoCliente.Equals(itemDescuento.TipoCliente));
                            if (tipoClienteDB != null)
                            {
                                descuentoPersistence.InsertDescuentoTramo(new DescuentoBE()
                                {
                                    idDescuento = Guid.NewGuid(),
                                    idTramo = idNuevoTramo,
                                    idTipoCliente = tipoClienteDB.idTipoCliente,
                                    DtoMax = itemDescuento.DtoMax,
                                    DtoMaxTDC = itemDescuento.DtoMaxTDC,
                                    AgrupacionTramo = itemDescuento.AgrupacionTramo
                                });
                            }
                        }
                    }

                    #endregion
                }

                //Guardamos las agrupaciones obtenidas para el producto
                agrupacionPersistence.GuardarAgrupaciones(listadoAgrupaciones, itemProductoAnterior.idProducto);

            }

            #endregion

            #region Características Producto

            foreach (ConfiguracionCaracteristicaBE caracteristica in listaCaracteristicas)
            {
                bool existeCaracteristica = false;

                //Obtengo la característica asociada
                caracteristica.idCaracteristica = caracteristicaPersistence.ObtenerIDCaracteristica(caracteristica.NombreCaracteristica, caracteristica.DescripcionCaracteristica, out existeCaracteristica);
                caracteristica.idProducto = idNuevoProducto;

                if (!existeCaracteristica)
                {
                    //Sólo inserto los valores de la característica si no existía previamente
                    foreach (var item in caracteristica.ListaValores)
                    {
                        item.idValor = caracteristicaPersistence.InsertUpdateValor(item.Valor, item.Descripcion, caracteristica.idCaracteristica);
                    }
                }
                else
                {
                    foreach (var item in caracteristica.ListaValores)
                    {
                        item.idValor = caracteristicaPersistence.ObtenerIdValor(caracteristica.idCaracteristica, item.Valor);
                    }
                }

                //cuando ya tengo todo el proceso anterior, añado un nuevo registro en bbdd
                caracteristicaPersistence.InsertCaracteristicaProducto(caracteristica);
            }

            #endregion

            #region Esqueleto VA

            Collection<EsquemaProductoBE> auxVAs = esqueletoProducto.Where(x => string.IsNullOrWhiteSpace(x.CodTramoSAP) && string.IsNullOrWhiteSpace(x.CodDestinoSAP) && !string.IsNullOrWhiteSpace(x.VASAP)).ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>();
            foreach (EsquemaProductoBE itemVA in auxVAs)
            {
                //solo se tratan los que realmente son VA...
                if ((!itemVA.VASAP.Equals("S0098")) && (!itemVA.VASAP.Equals("S0099")) && (!itemVA.VASAP.Equals("S0134")))
                {
                    #region Añadir VA al maestro de VA

                    ValorAnadidoBE valorAnadidoDB = listaVA.FirstOrDefault(x => x.CodValorAnadidoSAP.Equals(itemVA.VASAP));
                    if (valorAnadidoDB == null || String.IsNullOrEmpty(valorAnadidoDB.Descripcion))
                    {
                        Boolean isUpdate = valorAnadidoDB != null && String.IsNullOrEmpty(valorAnadidoDB.Descripcion);

                        //No está en base de datos, hay que añadirlo e insertarlo
                        if (valorAnadidoDB == null)
                        {
                            valorAnadidoDB = new ValorAnadidoBE()
                            {
                                idValorAnadido = Guid.NewGuid(),
                                CodValorAnadidoSAP = itemVA.VASAP,
                                Descripcion = itemVA.DescripcionVA
                            };
                        }
                        else
                        {
                            //Existe pero le faltan valores
                            valorAnadidoDB.Descripcion = itemVA.DescripcionVA;
                        }

                        // Lo rellenamos con FirstOrDefault porque en los preciospor del VA dado tanto
                        // negociable por precio cierto (true, false) como modalidad de negociacion (individual, general)
                        // tienen que ser las mismas para todos
                        PrecioPorBE precioPorVA = listaPrecioPor.FirstOrDefault(x => x.ValorAnadidoSAP.Equals(itemVA.VASAP));
                        if (precioPorVA != null)
                        {
                            valorAnadidoDB.EsParametrizable = true;
                            valorAnadidoDB.NegociableAPC = precioPorVA.NegociablePorPrecioCierto;
                            valorAnadidoDB.ModalidadNegociacion = precioPorVA.ModalidadNegociacionTarifa;
                        }

                        if (isUpdate)
                        {
                            vaPersistence.InsertUpdateValorAnadido(valorAnadidoDB);
                        }
                        else
                        {
                            vaPersistence.InsertValorAnadido(valorAnadidoDB);
                            listaVA.Add(valorAnadidoDB);

                        }
                    }

                    #endregion

                    #region Relación del VA con el producto

                    //inserto la relación del producto y del VA                    
                    ValorAnadidoAnexoProductoBE auxValorAnadidoProducto = new ValorAnadidoAnexoProductoBE()
                    {
                        idValorAnadidoAnexoProducto = Guid.NewGuid(),
                        idProducto = idNuevoProducto,
                        idValorAnadido = valorAnadidoDB.idValorAnadido
                    };
                    vapPersistence.InsertValorAnadidoProducto(auxValorAnadidoProducto);

                    #region Tipología del VA
                    //guardo la tipologia del VA
                    foreach (TipologiaClienteBE itemTipologia in listaTipologias.Where(x => !string.IsNullOrWhiteSpace(x.CodValorAnadidoSAP) && x.CodValorAnadidoSAP.Equals(itemVA.VASAP)))
                    {
                        TipoClienteBE tipoClienteDB = tiposClientesDB.FirstOrDefault(x => x.CodTipoCliente.Equals(itemTipologia.TipoCliente));
                        if (tipoClienteDB != null)
                        {
                            tipologiaPersistence.InsertTipologiaVA(new TipologiaClienteBE()
                            {
                                idTipologiaVA = Guid.NewGuid(),
                                idValorAnadidoProducto = auxValorAnadidoProducto.idValorAnadidoAnexoProducto,
                                idTipoCliente = tipoClienteDB.idTipoCliente,
                                Potencialidad = itemTipologia.Potencialidad,
                                Facturacion = itemTipologia.Facturacion
                            });
                        }

                    }

                    #endregion

                    #region Tarifa y descuento del VA

                    //tarifa VA
                    foreach (TarifaSAPBE obTarifa in listaTarifas.Where(x => x.VA.Equals(itemVA.VASAP) && !x.VA.Equals("SAC059")))
                    {
                        TarifaBE tarifaProductoDB = new TarifaBE();

                        tarifaProductoDB.idTarifa = Guid.NewGuid();
                        tarifaProductoDB.TipoPrecio = obTarifa.TipoPrecio;

                        decimal auxTarifa = 0;
                        decimal.TryParse(obTarifa.Tarifa.ToString(), out auxTarifa);

                        tarifaProductoDB.Tarifa = auxTarifa;
                        tarifaProductoDB.Moneda = obTarifa.Moneda;
                        tarifaProductoDB.idValorAnadidoProducto = auxValorAnadidoProducto.idValorAnadidoAnexoProducto;

                        PrecioPorBE precioPorVA = listaPrecioPor.FirstOrDefault(x => x.ValorAnadidoSAP.Equals(obTarifa.VA) && x.UnidadDeMedida.Equals(obTarifa.Tipopreciode));

                        if (precioPorVA != null)
                        {
                            tarifaProductoDB.Combinacion = obTarifa.Combinacion;
                            tarifaProductoDB.TipoTarifa = precioPorVA.TipoDePrecio;
                            tarifaProductoDB.NombreTarifa = precioPorVA.UnidadDeMedida;
                        }

                        tarifaPersistence.InsertTarifaVA(tarifaProductoDB);
                    }

                    //Caso especial para el SAC059 solo se debe insertar la primera tarifa
                    TarifaSAPBE objTarifa = listaTarifas.FirstOrDefault(x => x.VA.Equals("SAC059") && x.VA.Equals(itemVA.VASAP));
                    if (objTarifa != null)
                    {
                        TarifaBE tarifaDB = new TarifaBE();

                        tarifaDB.idTarifa = Guid.NewGuid();
                        tarifaDB.TipoPrecio = objTarifa.TipoPrecio;

                        decimal auxTarifaEspecial = 0;
                        decimal.TryParse(objTarifa.Tarifa.ToString(), out auxTarifaEspecial);

                        tarifaDB.Tarifa = auxTarifaEspecial;
                        tarifaDB.Moneda = objTarifa.Moneda;
                        tarifaDB.idValorAnadidoProducto = auxValorAnadidoProducto.idValorAnadidoAnexoProducto;

                        tarifaPersistence.InsertTarifaVA(tarifaDB);
                    }

                    //descuentos VA                    
                    foreach (DescuentoSAPBE itemDescuento in listaDescuento.Where(x => !string.IsNullOrWhiteSpace(x.VA) && x.VA.Equals(itemVA.VASAP)))
                    {
                        TipoClienteBE tipoClienteDB = tiposClientesDB.FirstOrDefault(x => x.CodTipoCliente.Equals(itemDescuento.TipoCliente));
                        if (tipoClienteDB != null)
                        {
                            descuentoPersistence.InsertDescuentoVA(new DescuentoBE()
                            {
                                idDescuento = Guid.NewGuid(),
                                idValorAnadidoProducto = auxValorAnadidoProducto.idValorAnadidoAnexoProducto,
                                idTipoCliente = tipoClienteDB.idTipoCliente,
                                DtoMax = itemDescuento.DtoMax,
                                DtoMaxTDC = itemDescuento.DtoMaxTDC
                            });
                        }
                    }

                    #endregion

                    ConfiguracionCaracteristicaBE caracteristica = listaCaracteristicas.FirstOrDefault(x => x.CodVASAP.Equals(itemVA.VASAP));
                    if (caracteristica != null)
                    {
                        foreach (ValoresBE valor in caracteristica.ListaValores)
                        {
                            vapPersistence.GuardarValorAnadidoValores(valor.idValor, valorAnadidoDB.idValorAnadido);
                        }
                    }

                    #endregion
                }
            }

            #endregion

            #region Relación Productos

            //Borramos las relaciones asociadas al producto
            relacionPersistence.DeleteRelacionesProdSAPSobrantes(itemProductoAnterior.CodProducto, listaRelacionProductos);

            //Insertamos las relaciones asociadas al producto
            var listaRelacionesProducto = listaRelacionProductos.Where(t => t.CodProductoSAP_A.Equals(itemProductoAnterior.CodProducto) || t.CodProductoSAP_B.Equals(itemProductoAnterior.CodProducto));

            foreach (var item in listaRelacionesProducto)
            {
                relacionPersistence.InsertRelacionProductos(item);
            }

            #endregion

            #endregion

        }


        /// <summary>
        /// Guarda la definición de un producto del modelo Volumetrico
        /// </summary>
        /// <param name="uow">contexto de la base de datos</param>
        /// <param name="itemProductoAnterior">producto del que se quiere actualizar la definicion</param>
        /// <param name="esqueletoProducto">esqueleto de los destinos tramos del producto descargado</param>
        /// <param name="listaDescuento">lista de descuentos del producto descargado</param>
        /// <param name="listaTarifas">lista de las tarifas del producto descargado</param>
        /// <param name="itemGrado">registro de definicion del grado del producto</param>
        /// <param name="listaPenalizaciones">lista de penalizaciones del producto</param>
        /// <param name="itemReglaridad">regularidad del producto que se actualiza</param>
        /// <param name="itemUmbralD2">umbral para el cual se debe mostrar el combo de población activa del producto</param>
        /// <param name="listaRangoPoblacion">lista de los rangos de población</param>
        /// <param name="listaVA">lista de los VA soportados en el sistema</param>
        private void GuardarDefinicionProductoVolumetrico(IUnitOfWork uow, ProductoBE itemProductoAnterior, Collection<EsquemaProductoBE> esqueletoProducto, Collection<DescuentoSAPBE> listaDescuento, Collection<TarifaSAPBE> listaTarifas, GradosBE itemGrado, Collection<PenalizacionRegularidadProductoBE> listaPenalizaciones, RegularidadBE itemReglaridad, UmbralBE itemUmbralD2, Collection<RangoPoblacionD2BE> listaRangoPoblacion, ref Collection<ValorAnadidoBE> listaVA, InternacionalBE itemInternacional, Collection<PrecioPorBE> listaPrecioPor, Collection<RelacionProductosBE> listaRelacionProductos, Collection<TipologiaClienteBE> listaTipologias)
        {
            #region variables usadas para el guardado

            ProductoPersistence productoPersistence = new ProductoPersistence(uow);
            TarifaPersistence tarifaPersistence = new TarifaPersistence(uow);
            DescuentoPersistence descuentoPersistence = new DescuentoPersistence(uow);
            DestinoPersistence destinoPersistence = new DestinoPersistence(uow);
            TramoPersistence tramoPersistence = new TramoPersistence(uow);
            ValorAnadidoPersistence vaPersistence = new ValorAnadidoPersistence(uow);
            ValorAnadidoProductoPersistence vapPersistence = new ValorAnadidoProductoPersistence(uow);
            GradoPersistence gradoPersistence = new GradoPersistence(uow);
            PenalizacionRegularidadProductoPersistence penalizacionPersistence = new PenalizacionRegularidadProductoPersistence(uow);
            RangoPoblacionD2Persistence rangoPoblacionD2Persistence = new RangoPoblacionD2Persistence(uow);
            RelacionProductosPersistence relacionPersistence = new RelacionProductosPersistence(uow);
            AgrupacionPersistence agrupacionPersistence = new AgrupacionPersistence(uow);


            //JCNS. DTO-MAX. incluyo listaTipologiasCliente
            TipoClientePersistence tipoClientePersistence = new TipoClientePersistence(uow);
            Collection<TipoClienteBE> tiposClientesDB = tipoClientePersistence.ObtenerTiposClientes();
            TipologiaClientePersistence tipologiaPersistence = new TipologiaClientePersistence(uow);

            #endregion

            #region proceso de guardado

            #region Producto

            //creo el nuevo registro en la tabla producto            
            Guid idNuevoProducto = Guid.NewGuid();

            ProductoBE nuevoProducto = new ProductoBE
            {
                idProducto = idNuevoProducto,
                idAnexoProducto = itemProductoAnterior.idAnexoProducto,
                ValidezDesde = DateTime.Now,
                Regularidad = null,
                ActualizacionPendiente = false,
                UmbralD2 = null,
                CodProducto = itemProductoAnterior.CodProducto,
                CodAnexoSAP = itemProductoAnterior.CodAnexoSAP,
            };

            if (itemUmbralD2 != null)
            {
                nuevoProducto.UmbralD2 = itemUmbralD2.Umbral;
            }

            if (itemReglaridad != null)
            {
                nuevoProducto.Regularidad = itemReglaridad.Regularidad;
            }

            if (itemInternacional != null)
            {
                nuevoProducto.Internacional = true;
            }

            //Se guarda como un registro nuevo.                
            productoPersistence.InsertProducto(nuevoProducto);

            #endregion

            #region Tipologia Producto
            //JCNS. DTO-MAX. incluyo listaTipologiasCliente. NO LO HACÍA EN VOLUMETRICO
            
            //guardo la tipología del producto nuevo
            foreach (TipologiaClienteBE itemTipologia in listaTipologias.Where(x => string.IsNullOrWhiteSpace(x.CodValorAnadidoSAP)))
            {
                TipoClienteBE tipoClienteDB = tiposClientesDB.FirstOrDefault(x => x.CodTipoCliente.Equals(itemTipologia.TipoCliente));
                if (tipoClienteDB != null)
                {
                    tipologiaPersistence.InsertTipologiaProducto(new TipologiaClienteBE()
                    {
                        idTipologiaProducto = Guid.NewGuid(),
                        idProducto = idNuevoProducto,
                        idTipoCliente = tipoClienteDB.idTipoCliente,
                        Potencialidad = itemTipologia.Potencialidad,
                        Facturacion = itemTipologia.Facturacion
                    });
                }
            }

            #endregion

            #region Grados, regularidad y Rango Población

            if (itemGrado != null)
            {
                gradoPersistence.InsertGrado(new GradosBE()
                {
                    idGrado = Guid.NewGuid(),
                    idProducto = idNuevoProducto,
                    G0 = itemGrado.G0,
                    G1 = itemGrado.G1,
                    G2 = itemGrado.G2
                });
            }

            foreach (PenalizacionRegularidadProductoBE itemPRP in listaPenalizaciones)
            {
                penalizacionPersistence.InsertPenalizacionRegularidadProducto(new PenalizacionRegularidadProductoBE()
                {
                    idPenalizacionRegularidadProducto = Guid.NewGuid(),
                    idProducto = idNuevoProducto,
                    RegularidadDesde = itemPRP.RegularidadDesde,
                    RegularidadHasta = itemPRP.RegularidadHasta,
                    Penalizacion = itemPRP.Penalizacion,
                    Bonificacion = itemPRP.Bonificacion
                });
            }

            foreach (RangoPoblacionD2BE itemRangoPoblacion in listaRangoPoblacion)
            {
                rangoPoblacionD2Persistence.InsertRangosPoblacionD2(new RangoPoblacionD2BE()
                {
                    idRangoPobalcionD2 = Guid.NewGuid(),
                    idProducto = idNuevoProducto,
                    Rango = itemRangoPoblacion.Rango,
                    Bonificacion = itemRangoPoblacion.Bonificacion
                });
            }


            #endregion

            #region Destinos y tramos
            //Borramos las agrupaciones del producto, para introducir las nuevas
            Collection<AgrupacionBE> listadoAgrupaciones = new Collection<AgrupacionBE>();
            //agrupacionPersistence.BorrarAgrupacionesProducto(itemProductoAnterior.idProducto);

            Collection<EsquemaProductoBE> auxDestinos = esqueletoProducto.Where(x => !x.CodDestinoSAP.Equals("S/D") && string.IsNullOrWhiteSpace(x.CodTramoSAP) && string.IsNullOrWhiteSpace(x.VASAP)).ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>();
            Collection<EsquemaProductoBE> auxTramos = esqueletoProducto.Where(x => !string.IsNullOrWhiteSpace(x.CodTramoSAP) && string.IsNullOrWhiteSpace(x.VASAP)).ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>();
            foreach (EsquemaProductoBE itemDestino in auxDestinos.OrderBy(x => x.Orden))
            {
                Guid idNuevoDestino = Guid.NewGuid();

                #region destino y descuento del destino

                //Guardamos el destino
                DestinoBE destinoDB = new DestinoBE()
                {
                    idDestino = idNuevoDestino,
                    CodDestinoSAP = itemDestino.CodDestinoSAP,
                    idProducto = idNuevoProducto,
                    Orden = itemDestino.Orden
                };
                destinoPersistence.InsertDestino(destinoDB);

                //Guardamos las agrupaciones del destino
                if (itemDestino.Agrupacion != null && itemDestino.Agrupacion.Length > 0)
                {
                    foreach (var nombreAgrupacion in itemDestino.Agrupacion)
                    {
                        var nuevaAgrupacion = listadoAgrupaciones.FirstOrDefault(t => t.Nombre.Equals(nombreAgrupacion));
                        Boolean anyadirAgrupacion = false;

                        //Si no existe en la lista, la añadimos
                        if (nuevaAgrupacion == null)
                        {
                            nuevaAgrupacion = new AgrupacionBE()
                            {
                                idAgrupacion = Guid.NewGuid(),
                                Nombre = nombreAgrupacion,
                                AgrupacionesDestino = new List<AgrupacionDestinoBE>()
                            };
                                                        
                            anyadirAgrupacion = true;
                        }

                        //Indicamos si es la agrupación por defecto
                        if (nuevaAgrupacion.Nombre.Contains("DEFAULT_") || nuevaAgrupacion.AgrupacionDefecto)
                        {
                            nuevaAgrupacion.AgrupacionDefecto = true;
                            nuevaAgrupacion.Nombre = nuevaAgrupacion.Nombre.Replace("DEFAULT_", String.Empty);
                        }
                        else
                        {
                            nuevaAgrupacion.AgrupacionDefecto = false;
                        }

                        if (anyadirAgrupacion)
                            listadoAgrupaciones.Add(nuevaAgrupacion);
                        
                        nuevaAgrupacion.AgrupacionesDestino.Add(new AgrupacionDestinoBE()
                        {
                            idAgrupacionDestino = Guid.NewGuid(),
                            idAgrupacion = nuevaAgrupacion.idAgrupacion,
                            idDestino = destinoDB.idDestino
                        });
                    }
                }

                //guardamos los descuentos del destino
                foreach (DescuentoSAPBE itemDescuento in listaDescuento.Where(x => x.Destino.Equals(itemDestino.CodDestinoSAP)))
                {
                    descuentoPersistence.InsertDescuentoVolumetrico(new DescuentoBE()
                    {
                        idDescuento = Guid.NewGuid(),
                        idDestino = idNuevoDestino,
                        DtoMax = itemDescuento.DtoMax,
                        DtoMaxTDC = itemDescuento.DtoMaxTDC,
                        VolumenSuperior = itemDescuento.VolumenSuperior,
                        VolumenInferior = itemDescuento.VolumenInferior
                    });
                }

                #endregion

                #region tramos y sus tarifas

                //Collection<EsquemaProductoBE> auxTramos = esqueletoProducto.Where(x => x.CodDestinoSAP.Equals(itemDestino.CodDestinoSAP) && !string.IsNullOrWhiteSpace(x.CodTramoSAP) && string.IsNullOrWhiteSpace(x.VASAP)).ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>();
                foreach (EsquemaProductoBE itemTramo in auxTramos.OrderBy(x => x.CodTramoSAP))
                {
                    //guardamos el tramo
                    Guid idNuevoTramo = Guid.NewGuid();
                    TramoBE tramoDB = new TramoBE()
                    {
                        idTramo = idNuevoTramo,
                        idDestino = idNuevoDestino,
                        CodTramo = itemTramo.CodTramoSAP,
                        Descripcion = itemTramo.DescripcionTramo
                    };
                    tramoPersistence.InsertTramo(tramoDB);

                    //guardamos la tarifa del tramo
                    TarifaSAPBE obTarifa = listaTarifas.FirstOrDefault(x => x.Destino.Equals(itemDestino.CodDestinoSAP) && x.Tramo.Equals(itemTramo.CodTramoSAP));

                    //Si no tiene tarifa, comprobamos si pertenece a una Zona Padre
                    if (obTarifa == null && !String.IsNullOrEmpty(itemDestino.CodDestinoZonaSAP))
                    {
                        obTarifa = listaTarifas.FirstOrDefault(x => x.Destino.Equals(itemDestino.CodDestinoZonaSAP) && x.Tramo.Equals(itemTramo.CodTramoSAP));
                    }

                    decimal auxTarifa = 0;
                    string tipoPrecio = "ZR50(EUR)";
                    string moneda = "EUR";

                    if (obTarifa != null)
                    {
                        decimal.TryParse(obTarifa.Tarifa.ToString(), out auxTarifa);
                        tipoPrecio = obTarifa.TipoPrecio;
                        moneda = obTarifa.Moneda;
                    }

                    tarifaPersistence.InsertTarifaProducto(new TarifaBE()
                    {
                        idTarifa = Guid.NewGuid(),
                        Tarifa = auxTarifa,
                        idTramo = idNuevoTramo,
                        TipoPrecio = tipoPrecio,
                        Moneda = moneda
                    });

                }

                #endregion
            }

            //Guardamos las agrupaciones obtenidas para el producto
            agrupacionPersistence.GuardarAgrupaciones(listadoAgrupaciones, itemProductoAnterior.idProducto);

            #endregion

            #region VA

            Collection<EsquemaProductoBE> auxVAs = esqueletoProducto.Where(x => string.IsNullOrWhiteSpace(x.CodTramoSAP) && string.IsNullOrWhiteSpace(x.CodDestinoSAP) && !string.IsNullOrWhiteSpace(x.VASAP)).ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>();
            foreach (EsquemaProductoBE itemVA in auxVAs)
            {
                //solo se tratan los que realmente son VA...
                if ((!itemVA.VASAP.Equals("S0098")) && (!itemVA.VASAP.Equals("S0099")) && (!itemVA.VASAP.Equals("S0134")))
                {
                    #region Añadir VA al maestro de VA

                    ValorAnadidoBE valorAnadidoDB = listaVA.FirstOrDefault(x => x.CodValorAnadidoSAP.Equals(itemVA.VASAP));
                    if (valorAnadidoDB == null || String.IsNullOrEmpty(valorAnadidoDB.Descripcion))
                    {
                        Boolean isUpdate = valorAnadidoDB != null && String.IsNullOrEmpty(valorAnadidoDB.Descripcion);

                        //No está en base de datos, hay que añadirlo e insertarlo
                        if (valorAnadidoDB == null)
                        {
                            valorAnadidoDB = new ValorAnadidoBE()
                            {
                                idValorAnadido = Guid.NewGuid(),
                                CodValorAnadidoSAP = itemVA.VASAP,
                                Descripcion = itemVA.DescripcionVA
                            };
                        }
                        else
                        {
                            //Existe pero le faltan valores
                            valorAnadidoDB.Descripcion = itemVA.DescripcionVA;
                        }

                        // Lo rellenamos con FirstOrDefault porque en los preciospor del VA dado tanto
                        // negociable por precio cierto (true, false) como modalidad de negociacion (individual, general)
                        // tienen que ser las mismas para todos
                        PrecioPorBE precioPorVA = listaPrecioPor.FirstOrDefault(x => x.ValorAnadidoSAP.Equals(itemVA.VASAP));
                        if (precioPorVA != null)
                        {
                            valorAnadidoDB.EsParametrizable = true;
                            valorAnadidoDB.NegociableAPC = precioPorVA.NegociablePorPrecioCierto;
                            valorAnadidoDB.ModalidadNegociacion = precioPorVA.ModalidadNegociacionTarifa;
                        }

                        if (isUpdate)
                        {
                            vaPersistence.InsertUpdateValorAnadido(valorAnadidoDB);
                        }
                        else
                        {
                            vaPersistence.InsertValorAnadido(valorAnadidoDB);
                            listaVA.Add(valorAnadidoDB);

                        }
                    }

                    #endregion

                    #region Relación del VA con el producto

                    //inserto la relación del producto y del VA                    
                    ValorAnadidoAnexoProductoBE auxValorAnadidoProducto = new ValorAnadidoAnexoProductoBE()
                    {
                        idValorAnadidoAnexoProducto = Guid.NewGuid(),
                        idProducto = idNuevoProducto,
                        idValorAnadido = valorAnadidoDB.idValorAnadido
                    };
                    vapPersistence.InsertValorAnadidoProducto(auxValorAnadidoProducto);

                    #region Tarifa del VA

                    //tarifa VA                    
                    foreach (TarifaSAPBE obTarifa in listaTarifas.Where(x => x.VA.Equals(itemVA.VASAP)))
                    {
                        TarifaBE tarifaProductoDB = new TarifaBE();

                        tarifaProductoDB.idTarifa = Guid.NewGuid();
                        tarifaProductoDB.TipoPrecio = obTarifa.TipoPrecio;

                        decimal auxTarifa = 0;
                        decimal.TryParse(obTarifa.Tarifa.ToString(), out auxTarifa);

                        tarifaProductoDB.Tarifa = auxTarifa;
                        tarifaProductoDB.Moneda = obTarifa.Moneda;
                        tarifaProductoDB.idValorAnadidoProducto = auxValorAnadidoProducto.idValorAnadidoAnexoProducto;

                        tarifaPersistence.InsertTarifaVA(tarifaProductoDB);
                    }

                    #endregion

                    #endregion
                }
            }

            #endregion

            #region Relación Productos

            //Borramos las relaciones asociadas al producto
            relacionPersistence.DeleteRelacionesProdSAPSobrantes(itemProductoAnterior.CodProducto, listaRelacionProductos);

            //Insertamos las relaciones asociadas al producto
            var listaRelacionesProducto = listaRelacionProductos.Where(t => t.CodProductoSAP_A.Equals(itemProductoAnterior.CodProducto) || t.CodProductoSAP_B.Equals(itemProductoAnterior.CodProducto));

            foreach (var item in listaRelacionesProducto)
            {
                relacionPersistence.InsertRelacionProductos(item);
            }

            #endregion

            #endregion
        }

        /// <summary>
        /// Guarda la definicion de un producto de publicaciones
        /// </summary>
        /// <param name="uow">Contesto de la base de datos</param>
        /// <param name="itemProductoAnterior">producto del que se quiere actualiza la definicion</param>
        /// <param name="esqueletoProducto">Esquema de los destinos/tramos/VA que tiene el producto</param>
        /// <param name="listaDescuento">lista de los descuentos del producto</param>
        /// <param name="listaTarifas">lista de las tarifas del producto</param>
        /// <param name="listaPuntos">item del punto del producto</param>
        /// <param name="listaGrados">item del grado del producto</param>
        /// <param name="listaPenalizaciones">lista de penalizaciones</param>
        /// <param name="ListaReglaridad">item de la regularidad del producto</param>
        /// <param name="listaVA">lista de los valores añadidos del sistema</param>
        private void GuardarDefinicionProductoPublicaciones(IUnitOfWork uow, ProductoBE itemProductoAnterior, Collection<EsquemaProductoBE> esqueletoProducto, Collection<DescuentoSAPBE> listaDescuento, Collection<TarifaSAPBE> listaTarifas, PuntosBE itemPunto, GradosBE itemGrado, Collection<PenalizacionRegularidadProductoBE> listaPenalizaciones, RegularidadBE itemReglaridad, ref Collection<ValorAnadidoBE> listaVA, InternacionalBE itemInternacional, Collection<PrecioPorBE> listaPrecioPor, Collection<RelacionProductosBE> listaRelacionProductos, Collection<ConfiguracionCaracteristicaBE> listaCaracteristicas, Collection<TipologiaClienteBE> listaTipologias)
        {
            //JCNS. DTO-MAX. incluyo listaTipologias
            //this.GuardarDefinicionProductoCostes(uow, itemProductoAnterior, esqueletoProducto, listaDescuento, listaTarifas, itemPunto, itemGrado, listaPenalizaciones, itemReglaridad, ref listaVA, itemInternacional, listaPrecioPor, listaRelacionProductos, listaCaracteristicas);
            this.GuardarDefinicionProductoCostes(uow, itemProductoAnterior, esqueletoProducto, listaDescuento, listaTarifas, itemPunto, itemGrado, listaPenalizaciones, itemReglaridad, ref listaVA, itemInternacional, listaPrecioPor, listaRelacionProductos, listaCaracteristicas, listaTipologias);
        }

        /// <summary>
        /// Guarda la definicion de un producto de libros
        /// </summary>
        /// <param name="uow">Contesto de la base de datos</param>
        /// <param name="itemProductoAnterior">producto del que se quiere actualiza la definicion</param>
        /// <param name="esqueletoProducto">Esquema de los destinos/tramos/VA que tiene el producto</param>
        /// <param name="listaDescuento">lista de los descuentos del producto</param>
        /// <param name="listaTarifas">lista de las tarifas del producto</param>
        /// <param name="listaPuntos">item del punto del producto</param>
        /// <param name="listaGrados">item del grado del producto</param>
        /// <param name="listaPenalizaciones">lista de penalizaciones</param>
        /// <param name="ListaReglaridad">item de la regularidad del producto</param>
        /// <param name="listaVA">lista de los valores añadidos del sistema</param>
        private void GuardarDefinicionProductoLibros(IUnitOfWork uow, ProductoBE itemProductoAnterior, Collection<EsquemaProductoBE> esqueletoProducto, Collection<DescuentoSAPBE> listaDescuento, Collection<TarifaSAPBE> listaTarifas, PuntosBE itemPunto, GradosBE itemGrado, Collection<PenalizacionRegularidadProductoBE> listaPenalizaciones, RegularidadBE itemReglaridad, ref Collection<ValorAnadidoBE> listaVA, InternacionalBE itemInternacional, Collection<PrecioPorBE> listaPrecioPor, Collection<RelacionProductosBE> listaRelacionProductos, Collection<ConfiguracionCaracteristicaBE> listaCaracteristicas, Collection<TipologiaClienteBE> listaTipologias)
        {
            //JCNS. DTO-MAX. incluyo listaTipologias
            //this.GuardarDefinicionProductoCostes(uow, itemProductoAnterior, esqueletoProducto, listaDescuento, listaTarifas, itemPunto, itemGrado, listaPenalizaciones, itemReglaridad, ref listaVA, itemInternacional, listaPrecioPor, listaRelacionProductos, listaCaracteristicas);
            this.GuardarDefinicionProductoCostes(uow, itemProductoAnterior, esqueletoProducto, listaDescuento, listaTarifas, itemPunto, itemGrado, listaPenalizaciones, itemReglaridad, ref listaVA, itemInternacional, listaPrecioPor, listaRelacionProductos, listaCaracteristicas, listaTipologias);
        }

        /// <summary>
        /// Guarda la definicion de un producto de publicorreo
        /// </summary>
        /// <param name="uow">Contesto de la base de datos</param>
        /// <param name="itemProductoAnterior">producto del que se quiere actualiza la definicion</param>
        /// <param name="esqueletoProducto">Esquema de los destinos/tramos/VA que tiene el producto</param>
        /// <param name="listaDescuento">lista de los descuentos del producto</param>
        /// <param name="listaTarifas">lista de las tarifas del producto</param>
        /// <param name="listaPuntos">item del punto del producto</param>
        /// <param name="listaGrados">item del grado del producto</param>
        /// <param name="listaPenalizaciones">lista de penalizaciones</param>
        /// <param name="ListaReglaridad">item de la regularidad del producto</param>
        /// <param name="listaVA">lista de los valores añadidos del sistema</param>
        private void GuardarDefinicionProductoPublicorreo(IUnitOfWork uow, ProductoBE itemProductoAnterior, Collection<EsquemaProductoBE> esqueletoProducto, Collection<DescuentoSAPBE> listaDescuento, Collection<TarifaSAPBE> listaTarifas, PuntosBE itemPunto, GradosBE itemGrado, Collection<PenalizacionRegularidadProductoBE> listaPenalizaciones, RegularidadBE itemReglaridad, ref Collection<ValorAnadidoBE> listaVA, InternacionalBE itemInternacional, Collection<PrecioPorBE> listaPrecioPor, Collection<RelacionProductosBE> listaRelacionProductos, Collection<ConfiguracionCaracteristicaBE> listaCaracteristicas, Collection<TipologiaClienteBE> listaTipologias)
        {
            //JCNS. DTO-MAX. incluyo listaTipologias
            //this.GuardarDefinicionProductoCostes(uow, itemProductoAnterior, esqueletoProducto, listaDescuento, listaTarifas, itemPunto, itemGrado, listaPenalizaciones, itemReglaridad, ref listaVA, itemInternacional, listaPrecioPor, listaRelacionProductos, listaCaracteristicas);
            this.GuardarDefinicionProductoCostes(uow, itemProductoAnterior, esqueletoProducto, listaDescuento, listaTarifas, itemPunto, itemGrado, listaPenalizaciones, itemReglaridad, ref listaVA, itemInternacional, listaPrecioPor, listaRelacionProductos, listaCaracteristicas, listaTipologias);
        }

        /// <summary>
        /// Guarda la definicion de un producto de costes
        /// </summary>
        /// <param name="uow">Contesto de la base de datos</param>
        /// <param name="itemProductoAnterior">producto del que se quiere actualiza la definicion</param>
        /// <param name="esqueletoProducto">Esquema de los destinos/tramos/VA que tiene el producto</param>
        /// <param name="listaDescuento">lista de los descuentos del producto</param>
        /// <param name="listaTarifas">lista de las tarifas del producto</param>
        /// <param name="listaPuntos">item del punto del producto</param>
        /// <param name="listaGrados">item del grado del producto</param>
        /// <param name="listaPenalizaciones">lista de penalizaciones</param>
        /// <param name="ListaReglaridad">item de la regularidad del producto</param>
        /// <param name="listaVA">lista de los valores añadidos del sistema</param>
        private void GuardarDefinicionProductoCostes(IUnitOfWork uow, ProductoBE itemProductoAnterior, Collection<EsquemaProductoBE> esqueletoProducto, Collection<DescuentoSAPBE> listaDescuento, Collection<TarifaSAPBE> listaTarifas, PuntosBE itemPunto, GradosBE itemGrado, Collection<PenalizacionRegularidadProductoBE> listaPenalizaciones, RegularidadBE itemReglaridad, ref Collection<ValorAnadidoBE> listaVA, InternacionalBE itemInternacional, Collection<PrecioPorBE> listaPrecioPor, Collection<RelacionProductosBE> listaRelacionProductos, Collection<ConfiguracionCaracteristicaBE> listaCaracteristicas, Collection<TipologiaClienteBE> listaTipologias)
        {
            #region variables usadas para el guardado

            ProductoPersistence productoPersistence = new ProductoPersistence(uow);
            TarifaPersistence tarifaPersistence = new TarifaPersistence(uow);
            DescuentoPersistence descuentoPersistence = new DescuentoPersistence(uow);
            DestinoPersistence destinoPersistence = new DestinoPersistence(uow);
            TramoPersistence tramoPersistence = new TramoPersistence(uow);
            ValorAnadidoPersistence vaPersistence = new ValorAnadidoPersistence(uow);
            ValorAnadidoProductoPersistence vapPersistence = new ValorAnadidoProductoPersistence(uow);
            GradoPersistence gradoPersistence = new GradoPersistence(uow);
            PuntoPersistence puntoPersistence = new PuntoPersistence(uow);
            PenalizacionRegularidadProductoPersistence penalizacionPersistence = new PenalizacionRegularidadProductoPersistence(uow);
            CaracteristicaPersistence caracteristicaPersistence = new CaracteristicaPersistence(uow);
            RelacionProductosPersistence relacionPersistence = new RelacionProductosPersistence(uow);
            AgrupacionPersistence agrupacionPersistence = new AgrupacionPersistence(uow);
            TipoClientePersistence tipoClientePersistence = new TipoClientePersistence(uow);
            Collection<TipoClienteBE> tiposClientesDB = tipoClientePersistence.ObtenerTiposClientes();

            
            //JCNS. DTO-MAX. incluyo listaTipologiasCliente
            TipologiaClientePersistence tipologiaPersistence = new TipologiaClientePersistence(uow);
            
            #endregion

            #region proceso de guardado

            #region Producto

            //creo el nuevo registro en la tabla producto            
            Guid idNuevoProducto = Guid.NewGuid();

            ProductoBE nuevoProducto = new ProductoBE
            {
                idProducto = idNuevoProducto,
                idAnexoProducto = itemProductoAnterior.idAnexoProducto,
                ValidezDesde = DateTime.Now,
                Regularidad = null,
                ActualizacionPendiente = false,
                UmbralD2 = null,
                CodProducto = itemProductoAnterior.CodProducto,
                CodAnexoSAP = itemProductoAnterior.CodAnexoSAP,
            };

            if (itemReglaridad != null)
            {
                nuevoProducto.Regularidad = itemReglaridad.Regularidad;
            }

            if (itemInternacional != null)
            {
                nuevoProducto.Internacional = true;
            }

            //Se guarda como un registro nuevo.                
            productoPersistence.InsertProducto(nuevoProducto);

            #endregion


            #region Tipologia Producto
            //JCNS. DTO-MAX. incluyo listaTipologiasCliente. NO LO HACÍA EN COSTES

            //guardo la tipología del producto nuevo
            foreach (TipologiaClienteBE itemTipologia in listaTipologias.Where(x => string.IsNullOrWhiteSpace(x.CodValorAnadidoSAP)))
            {
                TipoClienteBE tipoClienteDB = tiposClientesDB.FirstOrDefault(x => x.CodTipoCliente.Equals(itemTipologia.TipoCliente));
                if (tipoClienteDB != null)
                {
                    tipologiaPersistence.InsertTipologiaProducto(new TipologiaClienteBE()
                    {
                        idTipologiaProducto = Guid.NewGuid(),
                        idProducto = idNuevoProducto,
                        idTipoCliente = tipoClienteDB.idTipoCliente,
                        Potencialidad = itemTipologia.Potencialidad,
                        Facturacion = itemTipologia.Facturacion
                    });
                }
            }

            #endregion



            #region Grados, puntos y regularidad

            if (itemGrado != null)
            {
                gradoPersistence.InsertGrado(new GradosBE()
                {
                    idGrado = Guid.NewGuid(),
                    idProducto = idNuevoProducto,
                    G0 = itemGrado.G0,
                    G1 = itemGrado.G1,
                    G2 = itemGrado.G2
                });
            }

            if (itemPunto != null)
            {
                puntoPersistence.InsertPunto(new PuntosBE()
                {
                    idPunto = Guid.NewGuid(),
                    idProducto = idNuevoProducto,
                    CAM = itemPunto.CAM,
                    RUR = itemPunto.RUR,
                    URB = itemPunto.URB
                });
            }

            foreach (PenalizacionRegularidadProductoBE itemPRP in listaPenalizaciones)
            {
                penalizacionPersistence.InsertPenalizacionRegularidadProducto(new PenalizacionRegularidadProductoBE()
                {
                    idPenalizacionRegularidadProducto = Guid.NewGuid(),
                    idProducto = idNuevoProducto,
                    RegularidadDesde = itemPRP.RegularidadDesde,
                    RegularidadHasta = itemPRP.RegularidadHasta,
                    Penalizacion = itemPRP.Penalizacion,
                    Bonificacion = itemPRP.Bonificacion
                });

            }

            #endregion

            #region Destinos y tramos

            Collection<EsquemaProductoBE> auxDestinos = esqueletoProducto.Where(x => !x.CodDestinoSAP.Equals("S/D") && string.IsNullOrWhiteSpace(x.CodTramoSAP) && string.IsNullOrWhiteSpace(x.VASAP)).ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>();
            Collection<EsquemaProductoBE> auxTramos = esqueletoProducto.Where(x => !string.IsNullOrWhiteSpace(x.CodTramoSAP) && string.IsNullOrWhiteSpace(x.VASAP)).ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>();
            Collection<AgrupacionBE> agrupacionesProducto = new Collection<AgrupacionBE>();

            //Borramos las agrupaciones del producto, para introducir las nuevas
            Collection<AgrupacionBE> listadoAgrupaciones = new Collection<AgrupacionBE>();
            //agrupacionPersistence.BorrarAgrupacionesProducto(itemProductoAnterior.idProducto);

            bool sinFirma = false;

            if (auxDestinos.Count > 0)
            {

                foreach (EsquemaProductoBE itemDestino in auxDestinos.OrderBy(x => x.Orden))
                {
                    Guid idNuevoDestino = Guid.NewGuid();

                    #region destino y descuento del destino

                    //Guardamos el destino
                    DestinoBE destinoDB = new DestinoBE()
                    {
                        idDestino = idNuevoDestino,
                        CodDestinoSAP = itemDestino.CodDestinoSAP,
                        idProducto = idNuevoProducto,
                        Orden = itemDestino.Orden
                    };
                    destinoPersistence.InsertDestino(destinoDB);

                    //guardamos las agrupaciones del destino
                    if (itemDestino.Agrupacion != null && itemDestino.Agrupacion.Length > 0)
                    {
                        foreach (var nombreAgrupacion in itemDestino.Agrupacion)
                        {
                            var nuevaAgrupacion = listadoAgrupaciones.FirstOrDefault(t => t.Nombre.Equals(nombreAgrupacion.Replace("DEFAULT_", String.Empty)));
                            Boolean anyadirAgrupacion = false;

                            //Si no existe en la lista, la añadimos
                            //if ()
                            if (nuevaAgrupacion == null)
                            {
                                nuevaAgrupacion = new AgrupacionBE()
                                {
                                    idAgrupacion = Guid.NewGuid(),
                                    Nombre = nombreAgrupacion,
                                    AgrupacionesDestino = new List<AgrupacionDestinoBE>()
                                };

                                anyadirAgrupacion = true;
                            }

                            //Indicamos si es la agrupación por defecto
                            if (nuevaAgrupacion.Nombre.Contains("DEFAULT_") || nuevaAgrupacion.AgrupacionDefecto)
                            {
                                nuevaAgrupacion.AgrupacionDefecto = true;
                                nuevaAgrupacion.Nombre = nuevaAgrupacion.Nombre.Replace("DEFAULT_", String.Empty);
                            }
                            else
                            {
                                nuevaAgrupacion.AgrupacionDefecto = false;
                            }

                            if (anyadirAgrupacion)
                                listadoAgrupaciones.Add(nuevaAgrupacion);

                            nuevaAgrupacion.AgrupacionesDestino.Add(new AgrupacionDestinoBE()
                            {
                                idAgrupacionDestino = Guid.NewGuid(),
                                idAgrupacion = nuevaAgrupacion.idAgrupacion,
                                idDestino = destinoDB.idDestino
                            });
                        }
                    }

                    //guardamos los descuentos del destino
                    Collection<DescuentoSAPBE> listaDescuentos = listaDescuento.Where(x => string.IsNullOrWhiteSpace(x.VA) && x.Destino.Equals(itemDestino.CodDestinoSAP)).ToList().ToCollection();
                    if (listaDescuento != null)
                    {
                        foreach (DescuentoSAPBE itemDescuento in listaDescuentos)
                        {
                            TipoClienteBE tipoClienteDB = tiposClientesDB.FirstOrDefault(x => x.CodTipoCliente.Equals(itemDescuento.TipoCliente));
                            if (tipoClienteDB != null)
                            {
                                descuentoPersistence.InsertDescuentoProducto(new DescuentoBE()
                                {
                                    idDescuento = Guid.NewGuid(),
                                    idDestino = idNuevoDestino,
                                    idTipoCliente = tipoClienteDB.idTipoCliente,
                                    DtoMax = itemDescuento.DtoMax,
                                    DtoMaxTDC = itemDescuento.DtoMaxTDC

                                });
                            }
                        }
                    }

                    #endregion

                    #region tramos y sus tarifas

                    EsquemaProductoBE itemTramoOP = new EsquemaProductoBE();

                    //Collection<EsquemaProductoBE> auxTramos = esqueletoProducto.Where(x => x.CodDestinoSAP.Equals(itemDestino.CodDestinoSAP) && !string.IsNullOrWhiteSpace(x.CodTramoSAP) && string.IsNullOrWhiteSpace(x.VASAP)).ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>();
                    foreach (EsquemaProductoBE itemTramo in auxTramos.OrderBy(x => x.CodTramoSAP))
                    {

                        if (itemTramo.CodProductoSAP.Equals("S0012"))
                        {
                            itemTramoOP.CodTramoSAP = itemTramo.CodTramoSAP.PadLeft(10, '0');
                        }
                        else
                        {
                            itemTramoOP = itemTramo;
                        }

                        //guardamos el tramo
                        Guid idNuevoTramo = Guid.NewGuid();
                        TramoBE tramoDB = new TramoBE()
                        {
                            idTramo = idNuevoTramo,
                            idDestino = idNuevoDestino,
                            CodTramo = itemTramoOP.CodTramoSAP,
                            Descripcion = itemTramo.DescripcionTramo
                        };

                        tramoPersistence.InsertTramo(tramoDB);

                        //guardamos la tarifa del tramo
                        TarifaSAPBE obTarifa = listaTarifas.FirstOrDefault(x => x.Destino.Equals(itemDestino.CodDestinoSAP) && x.Tramo.Equals(itemTramo.CodTramoSAP));

                        //Si no tiene tarifa, comprobamos si pertenece a una Zona Padre
                        if (obTarifa == null && !String.IsNullOrEmpty(itemDestino.CodDestinoZonaSAP))
                        {
                            obTarifa = listaTarifas.FirstOrDefault(x => x.Destino.Equals(itemDestino.CodDestinoZonaSAP) && x.Tramo.Equals(itemTramo.CodTramoSAP));
                        }

                        decimal auxTarifa = 0;
                        string tipoPrecio = "ZR50(EUR)";
                        string moneda = "EUR";

                        if (obTarifa != null)
                        {
                            decimal.TryParse(obTarifa.Tarifa.ToString(), out auxTarifa);
                            tipoPrecio = obTarifa.TipoPrecio;
                            moneda = obTarifa.Moneda;
                        }

                        tarifaPersistence.InsertTarifaProducto(new TarifaBE()
                        {
                            idTarifa = Guid.NewGuid(),
                            Tarifa = auxTarifa,
                            idTramo = idNuevoTramo,
                            TipoPrecio = tipoPrecio,
                            Moneda = moneda
                        });

                        if (listaCaracteristicas != null && listaCaracteristicas.Any(x => x.DescripcionCaracteristica.Equals("Tipo de entrega")))
                        {

                            sinFirma = true;

                            string combinacion = String.Empty;

                            if (itemDestino.CodDestinoSAP == "ZA")
                            {
                                combinacion = "Z01" + "P" + itemTramo.CodTramoSAP;
                            }
                            else if (itemDestino.CodDestinoSAP == "ZB")
                            {
                                combinacion = "Z02" + "P" + itemTramo.CodTramoSAP;
                            }
                            else if (itemDestino.CodDestinoSAP == "ZC")
                            {
                                combinacion = "Z03" + "P" + itemTramo.CodTramoSAP;
                            }
                            else if (itemDestino.CodDestinoSAP == "ZD")
                            {
                                combinacion = "Z04" + "P" + itemTramo.CodTramoSAP;
                            }
                            else if (itemDestino.CodDestinoSAP == "ZE")
                            {
                                combinacion = "Z05" + "P" + itemTramo.CodTramoSAP;
                            }
                            else
                            {
                                combinacion = itemDestino.CodDestinoSAP + "P" + itemTramo.CodTramoSAP;
                            }

                            TarifaSAPBE tarifaSapProductoDBConFirma = listaTarifas.FirstOrDefault(x => x.Combinacion.EndsWith(combinacion));

                            if (tarifaSapProductoDBConFirma != null)
                            {

                                //Se guarda la tarifa
                                TarifaBE tarifaProductoDBConFirma = new TarifaBE();

                                tarifaProductoDBConFirma.idTarifa = Guid.NewGuid();
                                tarifaProductoDBConFirma.idProducto = idNuevoProducto;
                                tarifaProductoDBConFirma.TipoPrecio = tarifaSapProductoDBConFirma.TipoPrecio;

                                decimal auxTarifaCaracteristicaConFirma = 0;
                                decimal.TryParse(tarifaSapProductoDBConFirma.Tarifa.ToString(), out auxTarifaCaracteristicaConFirma);

                                tarifaProductoDBConFirma.Tarifa = auxTarifaCaracteristicaConFirma;
                                tarifaProductoDBConFirma.Moneda = tarifaSapProductoDBConFirma.Moneda;

                                tarifaProductoDBConFirma.Combinacion = tarifaSapProductoDBConFirma.Combinacion;

                                tarifaPersistence.InsertTarifaProductoSinDestino(tarifaProductoDBConFirma);
                            }

                            TarifaSAPBE tarifaSapProductoDBSinFirma = listaTarifas.FirstOrDefault(x => x.Combinacion.EndsWith(combinacion + "X"));

                            if (tarifaSapProductoDBSinFirma != null)
                            {

                                //Se guarda la tarifa
                                TarifaBE tarifaProductoDBSinFirma = new TarifaBE();

                                tarifaProductoDBSinFirma.idTarifa = Guid.NewGuid();
                                tarifaProductoDBSinFirma.idProducto = idNuevoProducto;
                                tarifaProductoDBSinFirma.TipoPrecio = tarifaSapProductoDBSinFirma.TipoPrecio;

                                decimal auxTarifaCaracteristicaSinFirma = 0;
                                decimal.TryParse(tarifaSapProductoDBSinFirma.Tarifa.ToString(), out auxTarifaCaracteristicaSinFirma);

                                tarifaProductoDBSinFirma.Tarifa = auxTarifaCaracteristicaSinFirma;
                                tarifaProductoDBSinFirma.Moneda = tarifaSapProductoDBSinFirma.Moneda;

                                tarifaProductoDBSinFirma.Combinacion = tarifaSapProductoDBSinFirma.Combinacion;

                                tarifaPersistence.InsertTarifaProductoSinDestino(tarifaProductoDBSinFirma);
                            }
                        }
                    }
                }
            }
                #endregion
            

            if (!sinFirma)
            {

                foreach (TarifaSAPBE objTarifa in listaTarifas.Where(x => string.IsNullOrWhiteSpace(x.VA) && !(string.IsNullOrEmpty(x.Combinacion))))
                {
                    //Se guarda la tarifa
                    TarifaBE tarifaProductoDB = new TarifaBE();

                    tarifaProductoDB.idTarifa = Guid.NewGuid();
                    tarifaProductoDB.idProducto = idNuevoProducto;
                    tarifaProductoDB.TipoPrecio = objTarifa.TipoPrecio;

                    decimal auxTarifa = 0;
                    decimal.TryParse(objTarifa.Tarifa.ToString(), out auxTarifa);

                    tarifaProductoDB.Tarifa = auxTarifa;
                    tarifaProductoDB.Moneda = objTarifa.Moneda;

                    tarifaProductoDB.Combinacion = objTarifa.Combinacion;

                    tarifaPersistence.InsertTarifaProductoSinDestino(tarifaProductoDB);
                }

                //TODO: Arreglar en SAP la descarga de este producto y que se guarde con el tipo de cliente que debe.
                if (itemProductoAnterior.CodProducto.Equals("S0367"))
                {

                    DescuentoSAPBE itemDescuento = listaDescuento.FirstOrDefault(x => string.IsNullOrWhiteSpace(x.VA));
                    if (itemDescuento != null)
                    {
                        TipoClienteBE tipoClienteDB = tiposClientesDB.FirstOrDefault(x => x.CodTipoCliente.Equals("E"));
                        if (tipoClienteDB != null)
                        {
                            descuentoPersistence.InsertDescuentoProductoSinDestino(new DescuentoBE()
                            {
                                idDescuento = Guid.NewGuid(),
                                idProducto = idNuevoProducto,
                                idTipoCliente = tipoClienteDB.idTipoCliente,
                                DtoMax = itemDescuento.DtoMax,
                                DtoMaxTDC = itemDescuento.DtoMaxTDC
                            });
                        }

                    }
                }
            }

            agrupacionPersistence.GuardarAgrupaciones(listadoAgrupaciones, itemProductoAnterior.idProducto);

            #endregion

            #region Características Producto

            foreach (ConfiguracionCaracteristicaBE caracteristica in listaCaracteristicas)
            {
                bool existeCaracteristica = false;

                //Obtengo la característica asociada
                caracteristica.idCaracteristica = caracteristicaPersistence.ObtenerIDCaracteristica(caracteristica.NombreCaracteristica, caracteristica.DescripcionCaracteristica, out existeCaracteristica);
                caracteristica.idProducto = idNuevoProducto;

                if (!existeCaracteristica)
                {
                    //Sólo inserto los valores de la característica si no existía previamente
                    foreach (var item in caracteristica.ListaValores)
                    {
                        item.idValor = caracteristicaPersistence.InsertUpdateValor(item.Valor, item.Descripcion, caracteristica.idCaracteristica);
                    }
                }
                else
                {
                    foreach (var item in caracteristica.ListaValores)
                    {
                        item.idValor = caracteristicaPersistence.ObtenerIdValor(caracteristica.idCaracteristica, item.Valor);
                    }
                }

                //cuando ya tengo todo el proceso anterior, añado un nuevo registro en bbdd
                caracteristicaPersistence.InsertCaracteristicaProducto(caracteristica);
            }

            #endregion

            #region VA

            Collection<EsquemaProductoBE> auxVAs = esqueletoProducto.Where(x => string.IsNullOrWhiteSpace(x.CodTramoSAP) && string.IsNullOrWhiteSpace(x.CodDestinoSAP) && !string.IsNullOrWhiteSpace(x.VASAP)).ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>();
            foreach (EsquemaProductoBE itemVA in auxVAs)
            {
                //solo se tratan los que realmente son VA...
                if ((!itemVA.VASAP.Equals("S0098")) && (!itemVA.VASAP.Equals("S0099")) && (!itemVA.VASAP.Equals("S0134")))
                {
                    #region Añadir VA al maestro de VA

                    ValorAnadidoBE valorAnadidoDB = listaVA.FirstOrDefault(x => x.CodValorAnadidoSAP.Equals(itemVA.VASAP));
                    if (valorAnadidoDB == null || String.IsNullOrEmpty(valorAnadidoDB.Descripcion))
                    {
                        Boolean isUpdate = valorAnadidoDB != null && String.IsNullOrEmpty(valorAnadidoDB.Descripcion);

                        //No está en base de datos, hay que añadirlo e insertarlo
                        if (valorAnadidoDB == null)
                        {
                            valorAnadidoDB = new ValorAnadidoBE()
                            {
                                idValorAnadido = Guid.NewGuid(),
                                CodValorAnadidoSAP = itemVA.VASAP,
                                Descripcion = itemVA.DescripcionVA
                            };
                        }
                        else
                        {
                            //Existe pero le faltan valores
                            valorAnadidoDB.Descripcion = itemVA.DescripcionVA;
                        }

                        // Lo rellenamos con FirstOrDefault porque en los preciospor del VA dado tanto
                        // negociable por precio cierto (true, false) como modalidad de negociacion (individual, general)
                        // tienen que ser las mismas para todos
                        PrecioPorBE precioPorVA = listaPrecioPor.FirstOrDefault(x => x.ValorAnadidoSAP.Equals(itemVA.VASAP));
                        if (precioPorVA != null)
                        {
                            valorAnadidoDB.EsParametrizable = true;
                            valorAnadidoDB.NegociableAPC = precioPorVA.NegociablePorPrecioCierto;
                            valorAnadidoDB.ModalidadNegociacion = precioPorVA.ModalidadNegociacionTarifa;
                        }

                        if (isUpdate)
                        {
                            vaPersistence.InsertUpdateValorAnadido(valorAnadidoDB);
                        }
                        else
                        {
                            vaPersistence.InsertValorAnadido(valorAnadidoDB);
                            listaVA.Add(valorAnadidoDB);

                        }
                    }

                    #endregion

                    #region Relación del VA con el producto

                    //inserto la relación del producto y del VA                    
                    ValorAnadidoAnexoProductoBE auxValorAnadidoProducto = new ValorAnadidoAnexoProductoBE()
                    {
                        idValorAnadidoAnexoProducto = Guid.NewGuid(),
                        idProducto = idNuevoProducto,
                        idValorAnadido = valorAnadidoDB.idValorAnadido
                    };
                    vapPersistence.InsertValorAnadidoProducto(auxValorAnadidoProducto);

                    #region Tarifa del VA

                    if (listaTarifas.FirstOrDefault(x => x.VA.Equals(itemVA.VASAP)) != null)
                    {

                        //tarifa VA                    
                        foreach (TarifaSAPBE obTarifa in listaTarifas.Where(x => x.VA.Equals(itemVA.VASAP)))
                        {
                            TarifaBE tarifaProductoDB = new TarifaBE();

                            tarifaProductoDB.idTarifa = Guid.NewGuid();
                            tarifaProductoDB.TipoPrecio = obTarifa.TipoPrecio;

                            decimal auxTarifa = 0;
                            decimal.TryParse(obTarifa.Tarifa.ToString(), out auxTarifa);

                            tarifaProductoDB.Tarifa = auxTarifa;
                            tarifaProductoDB.Moneda = obTarifa.Moneda;
                            tarifaProductoDB.idValorAnadidoProducto = auxValorAnadidoProducto.idValorAnadidoAnexoProducto;

                            tarifaPersistence.InsertTarifaVA(tarifaProductoDB);
                        }

                    }
                    else if(itemVA.VASAP.Equals("SAC184"))
                    {
                        TarifaBE tarifaProductoDB = new TarifaBE();

                        tarifaProductoDB.idTarifa = Guid.NewGuid();

                        decimal auxTarifa = 0;

                        tarifaProductoDB.Tarifa = auxTarifa;
                        tarifaProductoDB.Moneda = "EUR";
                        tarifaProductoDB.idValorAnadidoProducto = auxValorAnadidoProducto.idValorAnadidoAnexoProducto;

                        tarifaPersistence.InsertTarifaVA(tarifaProductoDB);


                    }
                    #endregion

                    #endregion
                }
            }

            #endregion

            #region Relación Productos

            //Borramos las relaciones asociadas al producto
            relacionPersistence.DeleteRelacionesProdSAPSobrantes(nuevoProducto.CodProducto, listaRelacionProductos);

            //Insertamos las relaciones asociadas al producto
            var listaRelacionesProducto = listaRelacionProductos.Where(t => t.CodProductoSAP_A.Equals(nuevoProducto.CodProducto) || t.CodProductoSAP_B.Equals(nuevoProducto.CodProducto));

            foreach (var item in listaRelacionesProducto)
            {
                relacionPersistence.InsertRelacionProductos(item);
            }

            #endregion

            #endregion
        }

        /// <summary>
        /// Guarda la definición de un producto del modelo de paquetería.
        /// </summary>
        /// <param name="uow">Contexto de la base de datos</param>
        /// <param name="itemProductoAnterior">Producto del que se quiere guardar la definicion</param>
        /// <param name="esqueletoProducto">Esquema de los nodos del producto</param>
        /// <param name="listaDescuento">Lista de los descuentos aplicables al producto</param>
        /// <param name="listaTarifas">Lista de las tarifas apicables al producto</param>
        /// <param name="listaTipologias">Lista de tipologías apicables al producto</param>
        /// <param name="listaVA">lista de valores añadidos soportados en el sistema</param>
        private void GuardarDefinicionProductoPaqueteria(IUnitOfWork uow, ProductoBE itemProductoAnterior, Collection<EsquemaProductoBE> esqueletoProducto, Collection<DescuentoSAPBE> listaDescuento, Collection<TarifaSAPBE> listaTarifas, Collection<TipologiaClienteBE> listaTipologias, ref Collection<ValorAnadidoBE> listaVA, InternacionalBE itemInternacional, Collection<PrecioPorBE> listaPrecioPor, Collection<RelacionProductosBE> listaRelacionProductos)
        {
            #region variables usadas para el guardado

            ProductoPersistence productoPersistence = new ProductoPersistence(uow);
            TipologiaClientePersistence tipologiaPersistence = new TipologiaClientePersistence(uow);
            TarifaPersistence tarifaPersistence = new TarifaPersistence(uow);
            DescuentoPersistence descuentoPersistence = new DescuentoPersistence(uow);
            DestinoPersistence destinoPersistence = new DestinoPersistence(uow);
            TramoPersistence tramoPersistence = new TramoPersistence(uow);
            ValorAnadidoPersistence vaPersistence = new ValorAnadidoPersistence(uow);
            ValorAnadidoProductoPersistence vapPersistence = new ValorAnadidoProductoPersistence(uow);
            TipoClientePersistence tipoClientePersistence = new TipoClientePersistence(uow);
            Collection<TipoClienteBE> tiposClientesDB = tipoClientePersistence.ObtenerTiposClientes();
            RelacionProductosPersistence relacionPersistence = new RelacionProductosPersistence(uow);
            AgrupacionPersistence agrupacionPersistence = new AgrupacionPersistence(uow);

            #endregion

            #region Proceso Guardar

            #region Producto

            //creo el nuevo registro en la tabla producto            
            Guid idNuevoProducto = Guid.NewGuid();

            //Se guarda como un registro nuevo.                
            productoPersistence.InsertProducto(new ProductoBE
            {
                idProducto = idNuevoProducto,
                idAnexoProducto = itemProductoAnterior.idAnexoProducto,
                ValidezDesde = DateTime.Now,
                Regularidad = null,
                ActualizacionPendiente = false,
                UmbralD2 = null,
                CodProducto = itemProductoAnterior.CodProducto,
                CodAnexoSAP = itemProductoAnterior.CodAnexoSAP,
                Internacional = itemInternacional != null ? true : false,
            });

            #endregion

            #region Tipologia Producto

            //guardo la tipología del producto nuevo
            foreach (TipologiaClienteBE itemTipologia in listaTipologias.Where(x => string.IsNullOrWhiteSpace(x.CodValorAnadidoSAP)))
            {
                TipoClienteBE tipoClienteDB = tiposClientesDB.FirstOrDefault(x => x.CodTipoCliente.Equals(itemTipologia.TipoCliente));
                if (tipoClienteDB != null)
                {
                    tipologiaPersistence.InsertTipologiaProducto(new TipologiaClienteBE()
                    {
                        idTipologiaProducto = Guid.NewGuid(),
                        idProducto = idNuevoProducto,
                        idTipoCliente = tipoClienteDB.idTipoCliente,
                        Potencialidad = itemTipologia.Potencialidad,
                        Facturacion = itemTipologia.Facturacion
                    });
                }
            }

            #endregion

            #region Esqueleto de Destinos y Tramos
            //recorro (en función de si tiene o no) los destinos y tramos del producto y guardo la información de las tarifas y los descuentos

            //Borramos las agrupaciones del producto, para introducir las nuevas
            Collection<AgrupacionBE> listadoAgrupaciones = new Collection<AgrupacionBE>();
            //agrupacionPersistence.BorrarAgrupacionesProducto(itemProductoAnterior.idProducto);

            //esqueleto producto
            Collection<EsquemaProductoBE> auxDestinos = esqueletoProducto.Where(x => !x.CodDestinoSAP.Equals("S/D") && string.IsNullOrWhiteSpace(x.CodTramoSAP) && string.IsNullOrWhiteSpace(x.VASAP)).ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>();
            if (auxDestinos.Count.Equals(0))
            {
                //es un producto sin destinos
                #region Tarifas

                TarifaSAPBE objTarifa = listaTarifas.FirstOrDefault(x => string.IsNullOrWhiteSpace(x.VA));
                if (objTarifa != null)
                {
                    //Se guarda la tarifa
                    TarifaBE tarifaProductoDB = new TarifaBE();

                    tarifaProductoDB.idTarifa = Guid.NewGuid();
                    tarifaProductoDB.idProducto = idNuevoProducto;
                    tarifaProductoDB.TipoPrecio = objTarifa.TipoPrecio;

                    decimal auxTarifa = 0;
                    decimal.TryParse(objTarifa.Tarifa.ToString(), out auxTarifa);

                    tarifaProductoDB.Tarifa = auxTarifa;
                    tarifaProductoDB.Moneda = objTarifa.Moneda;

                    tarifaPersistence.InsertTarifaProductoSinDestino(tarifaProductoDB);
                }

                #endregion

                #region Descuentos

                foreach (DescuentoSAPBE itemDescuento in listaDescuento.Where(x => string.IsNullOrWhiteSpace(x.VA)))
                {
                    TipoClienteBE tipoClienteDB = tiposClientesDB.FirstOrDefault(x => x.CodTipoCliente.Equals(itemDescuento.TipoCliente));
                    if (tipoClienteDB != null)
                    {
                        descuentoPersistence.InsertDescuentoProductoSinDestino(new DescuentoBE()
                        {
                            idDescuento = Guid.NewGuid(),
                            idProducto = idNuevoProducto,
                            idTipoCliente = tipoClienteDB.idTipoCliente,
                            DtoMax = itemDescuento.DtoMax,
                            DtoMaxTDC = itemDescuento.DtoMaxTDC
                        });
                    }

                }

                #endregion
            }
            else
            {
                //es un producto con destinos y tramos
                Collection<EsquemaProductoBE> auxTramos = esqueletoProducto.Where(x => !string.IsNullOrWhiteSpace(x.CodTramoSAP) && string.IsNullOrWhiteSpace(x.VASAP)).ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>();

                foreach (EsquemaProductoBE itemDestino in auxDestinos.OrderBy(x => x.Orden))
                {
                    Guid idNuevoDestino = Guid.NewGuid();

                    #region destino y descuento del destino

                    //Guardamos el destino
                    DestinoBE destinoDB = new DestinoBE()
                    {
                        idDestino = idNuevoDestino,
                        CodDestinoSAP = itemDestino.CodDestinoSAP,
                        idProducto = idNuevoProducto,
                        Orden = itemDestino.Orden
                    };
                    destinoPersistence.InsertDestino(destinoDB);

                    //Guardamos las agrupaciones del destino
                    if (itemDestino.Agrupacion != null && itemDestino.Agrupacion.Length > 0)
                    {
                        foreach (var nombreAgrupacion in itemDestino.Agrupacion)
                        {
                            var nuevaAgrupacion = listadoAgrupaciones.FirstOrDefault(t => t.Nombre.Equals(nombreAgrupacion));
                            Boolean anyadirAgrupacion = false;

                            //Si no existe en la lista, la añadimos
                            if (nuevaAgrupacion == null)
                            {
                                nuevaAgrupacion = new AgrupacionBE()
                                {
                                    idAgrupacion = Guid.NewGuid(),
                                    Nombre = nombreAgrupacion,
                                    AgrupacionesDestino = new List<AgrupacionDestinoBE>()
                                };

                                anyadirAgrupacion = true;
                            }

                            //Indicamos si es la agrupación por defecto
                            if (nuevaAgrupacion.Nombre.Contains("DEFAULT_") || nuevaAgrupacion.AgrupacionDefecto)
                            {
                                nuevaAgrupacion.AgrupacionDefecto = true;
                                nuevaAgrupacion.Nombre = nuevaAgrupacion.Nombre.Replace("DEFAULT_", String.Empty);
                            }
                            else
                            {
                                nuevaAgrupacion.AgrupacionDefecto = false;
                            }

                            if (anyadirAgrupacion)
                                listadoAgrupaciones.Add(nuevaAgrupacion);

                            nuevaAgrupacion.AgrupacionesDestino.Add(new AgrupacionDestinoBE()
                            {
                                idAgrupacionDestino = Guid.NewGuid(),
                                idAgrupacion = nuevaAgrupacion.idAgrupacion,
                                idDestino = destinoDB.idDestino
                            });
                        }
                    }

                    //guardamos los descuentos del destino
                    foreach (DescuentoSAPBE itemDescuento in listaDescuento.Where(x => string.IsNullOrWhiteSpace(x.VA) && x.Destino.Equals(itemDestino.CodDestinoSAP)))
                    {
                        TipoClienteBE tipoClienteDB = tiposClientesDB.FirstOrDefault(x => x.CodTipoCliente.Equals(itemDescuento.TipoCliente));
                        if (tipoClienteDB != null)
                        {
                            descuentoPersistence.InsertDescuentoProducto(new DescuentoBE()
                            {
                                idDescuento = Guid.NewGuid(),
                                idDestino = idNuevoDestino,
                                idTipoCliente = tipoClienteDB.idTipoCliente,
                                DtoMax = itemDescuento.DtoMax,
                                DtoMaxTDC = itemDescuento.DtoMaxTDC
                            });
                        }
                    }

                    #endregion

                    #region tramos y sus tarifas

                    //Collection<EsquemaProductoBE> auxTramos = esqueletoProducto.Where(x => x.CodDestinoSAP.Equals(itemDestino.CodDestinoSAP) && !string.IsNullOrWhiteSpace(x.CodTramoSAP) && string.IsNullOrWhiteSpace(x.VASAP)).ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>();
                    foreach (EsquemaProductoBE itemTramo in auxTramos.OrderBy(x => x.CodTramoSAP))
                    {
                        //guardamos el tramo
                        Guid idNuevoTramo = Guid.NewGuid();
                        TramoBE tramoDB = new TramoBE()
                        {
                            idTramo = idNuevoTramo,
                            idDestino = idNuevoDestino,
                            CodTramo = itemTramo.CodTramoSAP,
                            Descripcion = itemTramo.DescripcionTramo
                        };
                        tramoPersistence.InsertTramo(tramoDB);

                        //guardamos la tarifa del tramo
                        TarifaSAPBE obTarifa = listaTarifas.FirstOrDefault(x => x.Destino.Equals(itemDestino.CodDestinoSAP) && x.Tramo.Equals(itemTramo.CodTramoSAP));

                        //Si no tiene tarifa, comprobamos si pertenece a una Zona Padre
                        if (obTarifa == null && !String.IsNullOrEmpty(itemDestino.CodDestinoZonaSAP))
                        {
                            obTarifa = listaTarifas.FirstOrDefault(x => x.Destino.Equals(itemDestino.CodDestinoZonaSAP) && x.Tramo.Equals(itemTramo.CodTramoSAP));
                        }

                        decimal auxTarifa = 0;
                        string tipoPrecio = "ZR50(EUR)";
                        string moneda = "EUR";

                        if (obTarifa != null)
                        {
                            decimal.TryParse(obTarifa.Tarifa.ToString(), out auxTarifa);
                            tipoPrecio = obTarifa.TipoPrecio;
                            moneda = obTarifa.Moneda;
                        }

                        tarifaPersistence.InsertTarifaProducto(new TarifaBE()
                        {
                            idTarifa = Guid.NewGuid(),
                            Tarifa = auxTarifa,
                            idTramo = idNuevoTramo,
                            TipoPrecio = tipoPrecio,
                            Moneda = moneda
                        });

                    }

                    #endregion
                }
                
                //Guardamos las agrupaciones obtenidas para el producto
                agrupacionPersistence.GuardarAgrupaciones(listadoAgrupaciones,itemProductoAnterior.idProducto);
            }

            #endregion

            #region Esqueleto VA

            Collection<EsquemaProductoBE> auxVAs = esqueletoProducto.Where(x => string.IsNullOrWhiteSpace(x.CodTramoSAP) && string.IsNullOrWhiteSpace(x.CodDestinoSAP) && !string.IsNullOrWhiteSpace(x.VASAP)).ToList<EsquemaProductoBE>().ToCollection<EsquemaProductoBE>();
            foreach (EsquemaProductoBE itemVA in auxVAs)
            {
                //solo se tratan los que realmente son VA...
                if ((!itemVA.VASAP.Equals("S0098")) && (!itemVA.VASAP.Equals("S0099")) && (!itemVA.VASAP.Equals("S0134")))
                {
                    #region Añadir VA al maestro de VA

                    ValorAnadidoBE valorAnadidoDB = listaVA.FirstOrDefault(x => x.CodValorAnadidoSAP.Equals(itemVA.VASAP));
                    if (valorAnadidoDB == null || String.IsNullOrEmpty(valorAnadidoDB.Descripcion))
                    {
                        Boolean isUpdate = valorAnadidoDB != null && String.IsNullOrEmpty(valorAnadidoDB.Descripcion);

                        //No está en base de datos, hay que añadirlo e insertarlo
                        if (valorAnadidoDB == null)
                        {
                            valorAnadidoDB = new ValorAnadidoBE()
                            {
                                idValorAnadido = Guid.NewGuid(),
                                CodValorAnadidoSAP = itemVA.VASAP,
                                Descripcion = itemVA.DescripcionVA
                            };
                        }
                        else
                        {
                            //Existe pero le faltan valores
                            valorAnadidoDB.Descripcion = itemVA.DescripcionVA;
                        }

                        // Lo rellenamos con FirstOrDefault porque en los preciospor del VA dado tanto
                        // negociable por precio cierto (true, false) como modalidad de negociacion (individual, general)
                        // tienen que ser las mismas para todos
                        PrecioPorBE precioPorVA = listaPrecioPor.FirstOrDefault(x => x.ValorAnadidoSAP.Equals(itemVA.VASAP));
                        if (precioPorVA != null)
                        {
                            valorAnadidoDB.EsParametrizable = true;
                            valorAnadidoDB.NegociableAPC = precioPorVA.NegociablePorPrecioCierto;
                            valorAnadidoDB.ModalidadNegociacion = precioPorVA.ModalidadNegociacionTarifa;
                        }
                        if (isUpdate)
                        {
                            vaPersistence.InsertUpdateValorAnadido(valorAnadidoDB);
                        }
                        else
                        {
                            vaPersistence.InsertValorAnadido(valorAnadidoDB);
                            listaVA.Add(valorAnadidoDB);
                        }
                    }

                    #endregion

                    #region Relación del VA con el producto

                    //inserto la relación del producto y del VA                    
                    ValorAnadidoAnexoProductoBE auxValorAnadidoProducto = new ValorAnadidoAnexoProductoBE()
                    {
                        idValorAnadidoAnexoProducto = Guid.NewGuid(),
                        idProducto = idNuevoProducto,
                        idValorAnadido = valorAnadidoDB.idValorAnadido
                    };
                    vapPersistence.InsertValorAnadidoProducto(auxValorAnadidoProducto);

                    #region Tipología del VA
                    //guardo la tipologia del VA
                    foreach (TipologiaClienteBE itemTipologia in listaTipologias.Where(x => !string.IsNullOrWhiteSpace(x.CodValorAnadidoSAP) && x.CodValorAnadidoSAP.Equals(itemVA.VASAP)))
                    {
                        TipoClienteBE tipoClienteDB = tiposClientesDB.FirstOrDefault(x => x.CodTipoCliente.Equals(itemTipologia.TipoCliente));
                        if (tipoClienteDB != null)
                        {
                            tipologiaPersistence.InsertTipologiaVA(new TipologiaClienteBE()
                            {
                                idTipologiaVA = Guid.NewGuid(),
                                idValorAnadidoProducto = auxValorAnadidoProducto.idValorAnadidoAnexoProducto,
                                idTipoCliente = tipoClienteDB.idTipoCliente,
                                Potencialidad = itemTipologia.Potencialidad,
                                Facturacion = itemTipologia.Facturacion
                            });
                        }

                    }

                    #endregion

                    #region Tarifa y descuento del VA

                    //tarifa VA      
                    foreach (TarifaSAPBE obTarifa in listaTarifas.Where(x => x.VA.Equals(itemVA.VASAP)))
                    {
                        TarifaBE tarifaProductoDB = new TarifaBE();

                        tarifaProductoDB.idTarifa = Guid.NewGuid();
                        tarifaProductoDB.TipoPrecio = obTarifa.TipoPrecio;

                        decimal auxTarifa = 0;
                        decimal.TryParse(obTarifa.Tarifa.ToString(), out auxTarifa);

                        tarifaProductoDB.Tarifa = auxTarifa;
                        tarifaProductoDB.Moneda = obTarifa.Moneda;
                        tarifaProductoDB.idValorAnadidoProducto = auxValorAnadidoProducto.idValorAnadidoAnexoProducto;


                        //[MMUNOZ] No se obtiene toda la información necesaria de la tarifa para SAC175 ni SAC174
                        if (itemVA.VASAP.Equals("SAC175") || itemVA.VASAP.Equals("SAC174") || itemVA.VASAP.Equals("SAC185"))
                        {
                            PrecioPorBE precioPorVA = listaPrecioPor.FirstOrDefault(x => x.ValorAnadidoSAP.Equals(obTarifa.VA) && x.UnidadDeMedida.Equals(obTarifa.Tipopreciode));

                            if (precioPorVA != null)
                            {
                                if (!String.IsNullOrEmpty(obTarifa.Combinacion))
                                    tarifaProductoDB.Combinacion = obTarifa.Combinacion;

                                if (!String.IsNullOrEmpty(precioPorVA.UnidadDeMedida))
                                    tarifaProductoDB.NombreTarifa = precioPorVA.UnidadDeMedida;

                                tarifaProductoDB.TipoTarifa = precioPorVA.TipoDePrecio;
                            }
                        }

                        tarifaPersistence.InsertTarifaVA(tarifaProductoDB);
                    }

                    //descuentos VA                    
                    foreach (DescuentoSAPBE itemDescuento in listaDescuento.Where(x => !string.IsNullOrWhiteSpace(x.VA) && x.VA.Equals(itemVA.VASAP)))
                    {
                        TipoClienteBE tipoClienteDB = tiposClientesDB.FirstOrDefault(x => x.CodTipoCliente.Equals(itemDescuento.TipoCliente));
                        if (tipoClienteDB != null)
                        {
                            descuentoPersistence.InsertDescuentoVA(new DescuentoBE()
                            {
                                idDescuento = Guid.NewGuid(),
                                idValorAnadidoProducto = auxValorAnadidoProducto.idValorAnadidoAnexoProducto,
                                idTipoCliente = tipoClienteDB.idTipoCliente,
                                DtoMax = itemDescuento.DtoMax,
                                DtoMaxTDC = itemDescuento.DtoMaxTDC
                            });
                        }
                    }

                    #endregion

                    #endregion
                }
            }

            #endregion

            #region Relación Productos

            //Borramos las relaciones asociadas al producto
            relacionPersistence.DeleteRelacionesProdSAPSobrantes(itemProductoAnterior.CodProducto, listaRelacionProductos);

            //Insertamos las relaciones asociadas al producto
            var listaRelacionesProducto = listaRelacionProductos.Where(t =>
                                                                       t.CodProductoSAP_A.Equals(itemProductoAnterior.CodProducto) ||
                                                                       t.CodProductoSAP_B.Equals(itemProductoAnterior.CodProducto));

            foreach (var item in listaRelacionesProducto)
            {
                relacionPersistence.InsertRelacionProductos(item);
            }

            #endregion

            #endregion
        }

        /// <summary>
        /// Método que elimina los coeficientes de un producto y de los valores añadidos de un producto
        /// </summary>
        /// <param name="idProducto">Identificador del producto</param>
        /// <param name="usuario">Código del usuario</param>
        private void EliminarCoeficientesProducto(Guid idProducto, string usuario, IUnitOfWork uow)
        {
            CoeficienteBL bl = new CoeficienteBL();
            //Se eliminan los coeficientes del producto
            bl.EliminarCoeficientesProducto(idProducto, usuario, uow);
            //Se eliminan los coeficientes de los valores añadidos del producto
            bl.EliminarCoeficientesVA(idProducto, usuario, uow);
        }

        /// <summary>
        /// Método que elimina las potencialidades de un producto y de sus valores añadadidos
        /// </summary>
        /// <param name="idProducto">Identificador del producto</param>
        /// <param name="usuario">Código del usuario</param>
        /// <param name="uow">Contexto de la base de datos</param>
        private void EliminarPotencialidadesProducto(Guid idProducto, string usuario, IUnitOfWork uow)
        {
            PotencialidadBL pbl = new PotencialidadBL();
            //Se eliminan las potencialidades del producto
            pbl.EliminarPotencialidadProducto(idProducto, usuario, uow);
            //Se eliminan las potencialidades de los valores añadidos del producto
            pbl.EliminarPotencialidadVA(idProducto, usuario, uow);
        }

        #endregion

    }
}
