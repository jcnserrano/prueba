using Correos.SimuladorOfertas.Common;
using Correos.SimuladorOfertas.DTOs;
using Correos.SimuladorOfertas.Persistence;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;

namespace Correos.SimuladorOfertas.Business
{
    public class ProductoBL
    {
        #region Métodos Obtener

        /// <summary>
        /// Obtiene el descuento Máximo del destino
        /// </summary>
        /// <param name="idDestino"></param>
        /// <returns></returns>
        public double ObtenerDescuentoMaxProducto(Guid idDestino)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                DescuentoPersistence descuentopersistence = new DescuentoPersistence(uow);
                return descuentopersistence.ObtenerDescuentoDestino(idDestino);
            }
        }

        /// <summary>
        /// Obtiene la tarifa del producto sin tarifa
        /// </summary>
        /// <param name="idProducto"></param>
        /// <returns></returns>
        public decimal ObtenerTarifaProductoSinDestino(Guid idProducto, string combinacionCaracteristicas)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProductoPersistence persistence = new ProductoPersistence(uow);
                return persistence.ObtenerTarifaProductoSinDestino(idProducto, combinacionCaracteristicas);
            }
        }

        /// <summary>
        /// Obtiene la tarifa del producto sin tarifa
        /// </summary>
        /// <param name="idProducto"></param>
        /// <returns></returns>
        //JCNS. TARIFAS S0410. 2020-11-05 INC000050616447. SE AÑADE parámetro idProducto
        public Collection<TarifaBE> ObtenerListaTarifaProductoSinDestino(string combinacionProducto, Guid idProducto)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProductoPersistence persistence = new ProductoPersistence(uow);
                return persistence.ObtenerListaTarifaProductoSinDestino(combinacionProducto, idProducto);
            }
        }



        /// <summary>
        /// Obtiene un producto a partir de su código
        /// </summary>
        /// <param name="codigoProducto">Código del producto</param>
        /// <returns>Entidad Producto</returns>
        public ProductoBE ObtenerProducto(string codigoProducto)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProductoBE producto = new ProductoBE();
                ProductoPersistence productoPersistence = new ProductoPersistence(uow);
                producto = productoPersistence.ObtenerProducto(codigoProducto);
                return producto;
            }
        }

        /// <summary>
        /// JCNS. No se si con un codigo SAP de producto puedo obtener más de un idProducto. Por eso hago esta
        /// Obtiene una de productos a partir de su código. 
        /// </summary>
        /// <param name="codigoProducto">Código del producto</param>
        /// <returns>Entidad Producto</returns>        
        public Collection<ProductoBE> ObtenerProductoLista(string codigoProducto)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                Collection<ProductoBE> productosDB = new Collection<ProductoBE>();

                ProductoPersistence productoPersistence = new ProductoPersistence(uow);
                productosDB = productoPersistence.ObtenerProductosByCodProductoSAP(codigoProducto);
                return productosDB;
            }
        }

        public Collection<ProductoBE> ObtenerListaProductosNuevos(string productosNuevos)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                Collection<ProductoBE> productosDB = new Collection<ProductoBE>();

                ProductoPersistence productoPersistence = new ProductoPersistence(uow);
                productosDB = productoPersistence.Descarga_Productos_Nuevos(productosNuevos);
                return productosDB;
            }
        }



        /// <summary>
        /// Obtiene el ID de la última definición del producto.
        /// </summary>
        /// <param name="codProductoSAP">Código SAP del producto</param>
        /// <param name="anexoProducto">Código SAP del anexo</param>
        /// <returns></returns>
        public Guid ObtenerGuidUltimaDefincionProducto(string codProductoSAP, String anexoProducto)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                Guid idProducto;
                ProductoPersistence productoPersistence = new ProductoPersistence(uow);
                idProducto = productoPersistence.ObtenerGuidUltimaDefincionProducto(codProductoSAP, anexoProducto);
                return idProducto;
            }
        }

        public bool ObtenerVisibilidadProducto(string codProductoSAP)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProductoPersistence productoPersistence = new ProductoPersistence(uow);
                return productoPersistence.ObtenerVisibilidadProducto(codProductoSAP);
            }
        }



        /// <summary>
        /// Método que devuelve un producto por su idProducto
        /// </summary>
        /// <param name="idProducto">Identificador del producto</param>
        /// <returns>Entidad ProductoBE</returns>
        public ProductoBE ObtenerProductoByIdproducto(Guid idProducto)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                return ObtenerProductoByIdproducto(idProducto, uow);
            }
        }

        /// <summary>
        /// Método que devuelve un producto por su idProducto
        /// </summary>
        /// <param name="idProducto">Identificador del producto</param>
        /// <returns>Entidad ProductoBE</returns>
        public ProductoBE ObtenerProductoByIdproducto(Guid idProducto, IUnitOfWork uow)
        {
            ProductoPersistence productoPersistence = new ProductoPersistence(uow);
            return productoPersistence.ObtenerProductoByIdproducto(idProducto);
        }

        /// <summary>
        /// Método que obtiene los destinos para el producto pasado por parámetro
        /// </summary>
        /// <param name="idProducto">Identificador del producto</param>
        /// <returns>Lista de DestinoBE</returns>
        public Collection<DestinoBE> ObtenerDestinosProducto(Guid idProducto)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                return ObtenerDestinosProducto(idProducto, uow);
            }
        }

        /// <summary>
        /// Método que obtiene los destinos para el producto pasado por parámetro
        /// </summary>
        /// <param name="idProducto">Identificador del producto</param>
        /// <returns>Lista de DestinoBE</returns>
        public Collection<DestinoBE> ObtenerDestinosProducto(Guid idProducto, IUnitOfWork uow)
        {
            DestinoPersistence destinoPersistence = new DestinoPersistence(uow);
            return destinoPersistence.ObtenerDestinosProducto(idProducto, uow);
        }

        /// <summary>
        /// Método que obtiene un producto a partir de su identificador
        /// </summary>
        /// <param name="idProducto"></param>
        /// <returns></returns>
        public ProductoBE ObtenerProductoExcel(Guid idProducto, Guid idOferta, Guid idModalidad, string codProductoSAP, List<DestinoBE> destinosVisibles)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProductoBE producto = new ProductoBE();

                //Primeramente se obtiene el producto
                ProductoPersistence productoPersistence = new ProductoPersistence(uow);
                producto = productoPersistence.ObtenerProducto(idProducto);

                if (producto != null)
                {
                    //A continuación se obtienen los destinos del producto
                    DestinoPersistence destinoPersistence = new DestinoPersistence(uow);
                    producto.Destinos = destinoPersistence.ObtenerDestinosProducto(idProducto, idOferta, idModalidad, codProductoSAP, destinosVisibles);

                    //Se obtienen los valores añadidos del producto
                    ValorAnadidoProductoPersistence vaPersistence = new ValorAnadidoProductoPersistence(uow);
                    producto.ValoresAnadidos = vaPersistence.ObtenerValoresAnadidosProducto(idProducto);

                    //Se obtiene la lista de precios del producto
                    ListaPreciosPersistence preciosPersistence = new ListaPreciosPersistence(uow);
                    producto.ListaPrecios = preciosPersistence.ObtenerListaPreciosProducto(idProducto, idOferta, idModalidad);
                }

                return producto;
            }
        }

        /// <summary>
        /// Método que obtiene el listado completo de productos del generador de ofertas
        /// </summary>
        /// <returns>Listado completo de productos</returns>
        public Collection<ProductoBE> ObtenerListadoProductos()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProductoPersistence productoPersistence = new ProductoPersistence(uow);
                return productoPersistence.ObtenerListadoProductos();
            }
        }

        /// <summary>
        /// Método que obtiene una colección de productos que tienen pendiente una descarga de la definición
        /// </summary>                
        /// <returns>Lista de productos que deben actualizarse</returns>
        public Collection<ProductoBE> MarcarListadoProductosParaDescargaDefinicion(Collection<ProductoBE> listaProductos)
        {
            Collection<ProductoBE> productosPorActualizar = null;
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProductoPersistence productoPersistence = new ProductoPersistence(uow);
                productosPorActualizar = productoPersistence.MarcarListadoProductosParaDescargaDefinicion(listaProductos);
                uow.Save();
                return productosPorActualizar;
            }
        }

        /// <summary>
        /// Método que obtiene un punto a partir del identificador de producto pasado por parámetro
        /// </summary>
        /// <param name="idProducto">Identificador del producto</param>
        /// <returns>Entidad PuntosBE</returns>
        public PuntosBE ObtenerPunto(Guid idProducto)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                return ObtenerPunto(idProducto, uow);
            }
        }

        /// <summary>
        /// Método que obtiene un punto a partir del identificador de producto pasado por parámetro
        /// </summary>
        /// <param name="idProducto">Identificador del producto</param>
        /// <returns>Entidad PuntosBE</returns>
        public PuntosBE ObtenerPunto(Guid idProducto, IUnitOfWork uow)
        {
            PuntoPersistence puntoPersistence = new PuntoPersistence(uow);
            return puntoPersistence.ObtenerPunto(idProducto);
            }

        /// <summary>
        /// Método que obtiene la configuración grado de la oferta
        /// </summary>
        /// <param name="idProductoOferta">Identificador del productoOferta</param>
        /// <returns>Entidad ConfiguracionGradoOfertaBE</returns>
        public ConfiguracionGradoOfertaBE ObtenerConfiguracionGradoOferta(Guid idProductoOferta)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ConfiguracionGradoOfertaPersistence configGradoOfertaPersistence = new ConfiguracionGradoOfertaPersistence(uow);
                return configGradoOfertaPersistence.ObtenerConfiguracionGradoOferta(idProductoOferta);
            }
        }

        /// <summary>
        /// Método que obtiene todos los registros de la tabla configuraciónGradoOferta pertenecientes a los productosOferta pasados por parámetro.
        /// </summary>
        /// <param name="listaIdsProductosOferta">Listado de identificadores de productosOferta</param>
        /// <returns>Colección de entidades ConfiguracionGradoOfertaBE</returns>
        public Collection<ConfiguracionGradoOfertaBE> ObtenerConfiguracionGradoOferta(Collection<Guid> listaIdsProductosOferta)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ConfiguracionGradoOfertaPersistence configGradoOfertaPersistence = new ConfiguracionGradoOfertaPersistence(uow);
                return configGradoOfertaPersistence.ObtenerConfiguracionGradoOferta(listaIdsProductosOferta);
            }
        }

        /// <summary>
        /// Método que copia la configuración de grados de un producto oferta
        /// </summary>
        /// /// <param name="idProductoOfertaOrigen">Identificador del productoOferta origen</param>
        /// <param name="idProductoOfertaDestino">Identificador del producto Oferta destino</param>
        /// <returns>Colección de entidades ConfiguracionGradoOfertaBE</returns>        
        public Collection<ConfiguracionGradoOfertaBE> CopiarConfiguracionGradoProductoOferta(Guid idProductoOfertaOrigen, Guid idProductoOfertaDestino)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ConfiguracionGradoOfertaPersistence configGradoOfertaPersistence = new ConfiguracionGradoOfertaPersistence(uow);
                return configGradoOfertaPersistence.CopiarConfiguracionGradoProductoOferta(idProductoOfertaOrigen, idProductoOfertaDestino);
            }
        }

        /// <summary>
        /// Método que obtiene las penalizaciones para la regularidad de un producto
        /// </summary>
        /// <param name="idProducto">Identificador del producto</param>
        /// <returns>Entidad PenalizacionRegularidadProductoBE</returns>
        public Collection<PenalizacionRegularidadProductoBE> ObtenerPenalizacionRegularidadProducto(Guid idProducto)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                return ObtenerPenalizacionRegularidadProducto(idProducto, uow);
            }
        }

        /// <summary>
        /// Método que obtiene las penalizaciones para la regularidad de un producto
        /// </summary>
        /// <param name="idProducto">Identificador del producto</param>
        /// <returns>Entidad PenalizacionRegularidadProductoBE</returns>
        public Collection<PenalizacionRegularidadProductoBE> ObtenerPenalizacionRegularidadProducto(Guid idProducto, IUnitOfWork uow)
        {
            PenalizacionRegularidadProductoPersistence penalizacionPersistence = new PenalizacionRegularidadProductoPersistence(uow);
            return penalizacionPersistence.ObtenerPenalizacionRegularidadProducto(idProducto);
        }

        /// <summary>
        /// Método que obtiene el descuento para el idDestino e idTipoCliente pasados por parámetro
        /// </summary>
        /// <param name="idDestino">Identificador del destino</param>
        /// <returns>Entidad DescuentoBE</returns>
        public DescuentoBE ObtenerDescuentoDestinoByIdDestino(Guid idDestino)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                return ObtenerDescuentoDestinoByIdDestino(idDestino, uow);
                DescuentoPersistence descuentoPersistence = new DescuentoPersistence(uow);
                return descuentoPersistence.ObtenerDescuentoDestinoByIdDestino(idDestino);
            }
        }

        /// <summary>
        /// Método que obtiene el descuento para el idDestino e idTipoCliente pasados por parámetro
        /// </summary>
        /// <param name="idDestino">Identificador del destino</param>
        /// <returns>Entidad DescuentoBE</returns>
        public DescuentoBE ObtenerDescuentoDestinoByIdDestino(Guid idDestino, IUnitOfWork uow)
        {
            DescuentoPersistence descuentoPersistence = new DescuentoPersistence(uow);
            return descuentoPersistence.ObtenerDescuentoDestinoByIdDestino(idDestino);
        }

        /// <summary>
        /// Método que obtiene una colección de registros de la tabla descuentoProductoSinDestino para un producto
        /// </summary>
        /// <param name="idProducto">Identificador del producto</param>
        /// <returns>Colección de entidades DescuentoBE</returns>
        public Collection<DescuentoBE> ObtenerDescuentoProductoSinDestino(Guid idProducto)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                return ObtenerDescuentoProductoSinDestino(idProducto, uow);
            }
        }

        /// <summary>
        /// Método que obtiene una colección de registros de la tabla descuentoProductoSinDestino para un producto
        /// </summary>
        /// <param name="idProducto">Identificador del producto</param>
        /// <returns>Colección de entidades DescuentoBE</returns>
        public Collection<DescuentoBE> ObtenerDescuentoProductoSinDestino(Guid idProducto, IUnitOfWork uow)
        {
            DescuentoPersistence descuentoPersistence = new DescuentoPersistence(uow);
            return descuentoPersistence.ObtenerDescuentoProductoSinDestino(idProducto);
        }

        /// <summary>
        /// Método que obtiene la lista de tramos del destino pasado por parámetro almacenados en el contenedor local, ya que todavía no se ha
        /// hecho el commit de la transacción de inserción de tramos
        /// </summary>
        /// <param name="idDestino">Identificador del destino</param>
        /// <returns>Lista de TramoBE del destino</returns>
        public Collection<TramoBE> ObtenerTramosByIdDestino(Guid idDestino)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                return ObtenerTramosByIdDestino(idDestino, uow);
            }
        }

        /// <summary>
        /// Método que obtiene la lista de tramos del destino pasado por parámetro almacenados en el contenedor local, ya que todavía no se ha
        /// hecho el commit de la transacción de inserción de tramos
        /// </summary>
        /// <param name="idDestino">Identificador del destino</param>
        /// <returns>Lista de TramoBE del destino</returns>
        public Collection<TramoBE> ObtenerTramosByIdDestino(Guid idDestino, IUnitOfWork uow)
        {
            TramoPersistence tramoPersistence = new TramoPersistence(uow);
            return tramoPersistence.ObtenerTramosByIdDestino(idDestino);
        }

        /// <summary>
        /// <para>Indica si el producto tiene destinos con tramos de expediciones</para>
        /// <para>(Mejoras de impresión de ofertas)</para>
        /// </summary>
        /// <param name="idProducto"></param>
        /// <returns></returns>
        public Boolean TieneDestinosConTramosDeExpediciones(Guid idProducto) 
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                TramoPersistence tramoPersistence = new TramoPersistence(uow);
                return tramoPersistence.TieneDestinosConTramosDeExpediciones(idProducto);
            }
        }

        /// <summary>
        /// Método que obtiene las tarifas de la tabla TarifaProducto de un tramo
        /// </summary>
        /// <param name="idTramo">Identificador del tramo</param>
        /// <returns>Colección de entidades TarifaBE</returns>
        public Collection<TarifaBE> ObtenerTarifaProductoByIdTramo(Guid idTramo)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                return ObtenerTarifaProductoByIdTramo(idTramo, uow);
            }

        }

        public Collection<TarifaBE> ObtenerTarifaProductoByIdTramo(Guid idTramo, IUnitOfWork uow)
        {
            TarifaPersistence tarifaPersistence = new TarifaPersistence(uow);
            return tarifaPersistence.ObtenerTarifaProductoByIdTramo(idTramo, uow);
        }

        /// <summary>
        /// Método que obtiene la tarifa para el producto, del que no existen destinos, pasados por parámetro
        /// </summary>
        /// <param name="idProducto">Identificador del producto del que no existen destino</param>
        /// <returns>Entidad TarifaBE</returns>
        public TarifaBE ObtenerTarifaProductoSinDestino(Guid idProducto)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                return ObtenerTarifaProductoSinDestino(idProducto, uow);
            }
        }

        /// <summary>
        /// Método que obtiene la tarifa para el producto, del que no existen destinos, pasados por parámetro
        /// </summary>
        /// <param name="idProducto">Identificador del producto del que no existen destino</param>
        /// <returns>Entidad TarifaBE</returns>
        public TarifaBE ObtenerTarifaProductoSinDestino(Guid idProducto, IUnitOfWork uow)
        {
            TarifaPersistence tarifaPersistence = new TarifaPersistence(uow);
            return tarifaPersistence.ObtenerTarifaProductoSinDestino(idProducto);
        }

        /// <summary>
        /// Método que obtiene los registros de la tabla ValorAñadidoProducto de un producto
        /// </summary>
        /// <param name="idProducto">Identificador del producto</param>
        /// <returns>Colección de entidades ValorAnadidoAnexoProductoBE</returns>
        public Collection<ValorAnadidoAnexoProductoBE> ObtenerValorAnyadidoProducto(Guid idProducto, IUnitOfWork uow)
        {
            ValorAnadidoPersistence valorAnadidoPersistence = new ValorAnadidoPersistence(uow);
            return valorAnadidoPersistence.ObtenerValoresAnyadidoProducto(idProducto);
        }

        /// <summary>
        /// Método que obtiene los valores añadidos de un producto
        /// </summary>
        /// <param name="idValorAnyadido">Identificador del producto</param>
        /// <returns>Colección de entidades ValorAnadidoBE</returns>
        public Collection<ValorAnadidoBE> ObtenerValorAnyadidoByIdProducto(Guid idProducto)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                return ObtenerValorAnyadidoByIdProducto(idProducto, uow);
            }
        }

        /// <summary>
        /// Método que obtiene los valores añadidos de un producto
        /// </summary>
        /// <param name="idValorAnyadido">Identificador del producto</param>
        /// <returns>Colección de entidades ValorAnadidoBE</returns>
        public Collection<ValorAnadidoBE> ObtenerValorAnyadidoByIdProducto(Guid idProducto, IUnitOfWork uow)
        {
            ValorAnadidoPersistence valorAnyadidoPersistence = new ValorAnadidoPersistence(uow);
            return valorAnyadidoPersistence.ObtenerValorAnyadidoByIdProducto(idProducto);
        }

        /// <summary>
        /// Método que obtiene las tarifas de todos los valores añadidos de un producto
        /// </summary>
        /// <param name="idProducto">Identificador del producto</param>
        /// <returns>Colección de entidades TarifaBE</returns>
        public Collection<TarifaBE> ObtenerTarifasValorAnyadidoProducto(Guid idProducto)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                return ObtenerTarifasValorAnyadidoProducto(idProducto, uow);
            }
        }

        /// <summary>
        /// Método que obtiene las tarifas de todos los valores añadidos de un producto
        /// </summary>
        /// <param name="idProducto">Identificador del producto</param>
        /// <returns>Colección de entidades TarifaBE</returns>
        public Collection<TarifaBE> ObtenerTarifasValorAnyadidoProducto(Guid idProducto, IUnitOfWork uow)
        {
            TarifaPersistence tarifaPersistence = new TarifaPersistence(uow);
            return tarifaPersistence.ObtenerTarifasValorAnyadidoProducto(idProducto);
        }

        /// <summary>
        /// Método que obtiene todos los descuentos de los valores añadidos de un producto
        /// </summary>
        /// <param name="idProducto">Identificador del producto</param>
        /// <returns>Colección de entidades DescuentoBE</returns>
        public Collection<DescuentoBE> ObtenerDescuentosValorAnadidoProducto(Guid idProducto)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                return ObtenerDescuentosValorAnadidoProducto(idProducto, uow);
            }
        }

        /// <summary>
        /// Método que obtiene todos los descuentos de los valores añadidos de un producto
        /// </summary>
        /// <param name="idProducto">Identificador del producto</param>
        /// <returns>Colección de entidades DescuentoBE</returns>
        public Collection<DescuentoBE> ObtenerDescuentosValorAnadidoProducto(Guid idProducto, IUnitOfWork uow)
        {
            DescuentoPersistence descuentoPersistence = new DescuentoPersistence(uow);
            return descuentoPersistence.ObtenerDescuentosValorAnadidoProducto(idProducto);
        }

        /// <summary>
        /// Método que obtiene todos los productos de la BD
        /// </summary>
        /// <returns>Colección de entidades ProductoBE</returns>
        public Collection<ProductoBE> ObtenerProductos()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProductoPersistence productoPersistence = new ProductoPersistence(uow);
                return productoPersistence.ObtenerProductos();
            }
        }

        /// <summary>
        /// Método que obtiene todos los productos de la BD
        /// </summary>
        /// <returns>Colección de entidades ProductoBE</returns>
        public Collection<VisibilidadProductoBE> ObtenerProductosSoloLectura()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProductoPersistence productoPersistence = new ProductoPersistence(uow);
                return productoPersistence.ObtenerProductosSoloLectura();
            }
        }

        /// <summary>
        /// Obtiene la lista de ConfiguracionValorAnadidoTarifa de un listado de ConfiguracionValorAnadidoBE dado
        /// </summary>
        /// <param name="cvaBE"></param>
        /// <returns></returns>
        public Collection<ConfiguracionValorAnadidoTarifaBE> ObtenerListaConfiguracionValorAnadidoTarifa(Collection<ConfiguracionValorAnadidoBE> listaConfigVA)
        { 
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ConfiguracionValorAnadidoPersistence cVAPersistence = new ConfiguracionValorAnadidoPersistence(uow);
                return cVAPersistence.ObtenerListaConfiguracionValorAnadidoTarifa(listaConfigVA);
            }
        }

        /// <summary>
        /// Obtiene la lista de ConfiguracionValorAnadidoCaracteristica de los ConfiguracionValorAnadidoBE dado
        /// </summary>
        /// <param name="listaConfVABE">Colección de entidades ConfiguracionValorAnadidoBE</param>
        /// <returns>Colección de entidades ConfiguracionValorAnadidoCaracteristicaBE</returns>
        public Collection<ConfiguracionValorAnadidoCaracteristicaBE> ObtenerListaConfiguracionValorAnadidoCaracteristica(Collection<ConfiguracionValorAnadidoBE> listaConfVABE)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ConfiguracionValorAnadidoPersistence cVAPersistence = new ConfiguracionValorAnadidoPersistence(uow);
                return cVAPersistence.ObtenerListaConfiguracionValorAnadidoCaracteristica(listaConfVABE);
            }
        }

        /// <summary>
        /// Método que obtiene todos los registros de ConfiguracionValorAnadidoTarifa de los productos que se le pasa por parámetro
        /// </summary>
        /// <param name="listaIdsProductoOferta">Colección de identificadords de productos</param>
        /// <returns>Colección de entidades ConfiguracionValorAnadidoTarifa</returns>
        public Collection<ConfiguracionValorAnadidoTarifaBE> ObtenerConfiguracionesVAtarifaOferta(Collection<Guid> listaIdsProductoOferta)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ConfiguracionValorAnadidoPersistence cVAPersistence = new ConfiguracionValorAnadidoPersistence(uow);
                return cVAPersistence.ObtenerConfiguracionesVAtarifaOferta(listaIdsProductoOferta);
            }
        }

        /// <summary>
        /// Método que obtiene todos los registros de ConfiguaracionValorAnadidoCaracteristica de los productos que se le pasa por parámetro
        /// </summary>
        /// <param name="listaIdsProductoOferta">Colección de identificadores de productos</param>
        /// <returns>Colección de entidades ConfiguracionValorAnadidoTarifa</returns>
        public Collection<ConfiguracionValorAnadidoCaracteristicaBE> ObtenerConfiguracionesVACaracteristicaOferta(Collection<Guid> listaIdsProductoOferta)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ConfiguracionValorAnadidoPersistence cVAPersistence = new ConfiguracionValorAnadidoPersistence(uow);
                return cVAPersistence.ObtenerConfiguracionesVACaracteristicaOferta(listaIdsProductoOferta);
            }
        }

        #endregion

        #region Métodos Guardar

        /// <summary>
        /// Envía a SAP una lista de productos para guardarlos
        /// </summary>
        /// <param name="productos">Lista de productos que se quieren guardar</param>
        public void SincronizarProductosConSAP(Collection<ProductoBE> productos)
        {
            //TO-DO: En este caso, hay que llamar al servicio WCF 
        }

        /// <summary>
        /// Método que guarda en BD las definiciones de productos almacenadas en memoria
        /// </summary>
        /// <param name="listaProductos">Colección de productos</param>
        /// <param name="listaTipologiasProducto">Coleccion de tipologíasProducto</param>
        /// <param name="listaGrados">Colección de grados</param>
        /// <param name="listaPuntos">Colección de puntos</param>
        /// <param name="listaPenalizacionesProducto">Colección de penalizacionRegularidaProducto</param>
        /// <param name="listaDestinos">Colección de Destinos</param>
        /// <param name="listaDescuentosProductoConDestino">Colección de descuentoProducto</param>
        /// <param name="listaDescuentosProductoSinDestino">Colección de descuentoProductoSinDestino</param>
        /// <param name="listaTramos">Colección de tramos</param>
        /// <param name="listaDescuentosTramos">Colección de descuentoTramo</param>
        /// <param name="listaTarifasProductoTramo">Colección de TarifasProducto</param>
        /// <param name="listaTarifasProductoSinDestino">Colección de tarifaProductoSinDestino</param>
        /// <param name="listaValoresAnyadidos">Colección de valoresAnyadidos</param>
        /// <param name="listaTipologiasValorAnadido">Colección de tipologiasValorAnadido</param>
        /// <param name="listaTarifasVA">Colección de tarifaValorAnadido</param>
        /// <param name="listaDescuentosVA">Colección de descuentoValorAnadido</param>
        /// <param name="listaDescuentosVolumetrico">Colección de descuentoVolumétrico</param>
        /// <param name="listaCaracteristicas">Colección de Características</param>
        /// <param name="listaCaracteristicasProducto">Colección de CaracteristicasProducto</param>
        public void GuardarDefinicionesDeProductos(Collection<ProductoBE> listaProductos, Collection<TipologiaClienteBE> listaTipologiasProducto, Collection<GradoProductoInformacionBE> listaGrados,
                                                    Collection<PuntosBE> listaPuntos, Collection<PenalizacionRegularidadProductoBE> listaPenalizacionesProducto, Collection<DestinoBE> listaDestinos,
                                                    Collection<DescuentoBE> listaDescuentosProductoConDestino, Collection<DescuentoBE> listaDescuentosProductoSinDestino, Collection<TramoBE> listaTramos,
                                                    Collection<DescuentoBE> listaDescuentosTramos, Collection<TarifaBE> listaTarifasProductoTramo, Collection<TarifaBE> listaTarifasProductoSinDestino,
                                                    Collection<ValorAnadidoBE> listaValoresAnyadidos, Collection<ValorAnadidoAnexoProductoBE> listaValorAnyadidoProducto, Collection<TipologiaClienteBE> listaTipologiasValorAnadido, Collection<TarifaBE> listaTarifasVA,
                                                    Collection<DescuentoBE> listaDescuentosVA, Collection<DescuentoBE> listaDescuentosVolumetrico, Collection<ConfiguracionCaracteristicaBE> listaCaracteristicas,
                                                    Collection<ConfiguracionCaracteristicaBE> listaCaracteristicasProducto, Collection<RangoPoblacionD2BE> listaRangoPoblacionD2, Collection<CaracteristicaBE> listaCaracteristicasValoresAnadidos)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                //Se guardan las definiciones de cada producto en BD
                foreach (ProductoBE producto in listaProductos)
                {
                    #region variables

                    ProductoPersistence productoPersistence = new ProductoPersistence(uow);
                    TipologiaClientePersistence tipologiaClientePersistence = new TipologiaClientePersistence(uow);
                    GradoPersistence gradoPersistence = new GradoPersistence(uow);
                    PuntoPersistence puntoPersistence = new PuntoPersistence(uow);
                    PenalizacionRegularidadProductoPersistence penalizacionProductoPersistence = new PenalizacionRegularidadProductoPersistence(uow);
                    DestinoPersistence destinoPersistence = new DestinoPersistence(uow);
                    DescuentoPersistence descuentoPersistence = new DescuentoPersistence(uow);
                    TramoPersistence tramoPersistence = new TramoPersistence(uow);
                    TarifaPersistence tarifaPersistence = new TarifaPersistence(uow);
                    ValorAnadidoPersistence valorAnadidoPersistence = new ValorAnadidoPersistence(uow);
                    CaracteristicaPersistence caracteristicasPersistence = new CaracteristicaPersistence(uow);

                    #endregion

                    #region GuardarProducto

                    //Antes de insertar el producto que no existe en la BD de excelWorbook, se debe comprobar si este producto tiene validezHasta a NULL
                    //en ese caso se debe rellenar la fecha validez hasta al producto antiguo si lo hubiera
                    if (producto.ValidezHasta.Equals(null))
                    {
                        productoPersistence.MarcarProductoObsoleto(producto);
                    }

                    productoPersistence.InsertProducto(producto);

                    #endregion

                    #region Guardar tipologías del producto

                    foreach (TipologiaClienteBE tipologiaProducto in listaTipologiasProducto.Where(x => x.idProducto.Equals(producto.idProducto)))
                    {
                        tipologiaClientePersistence.InsertTipologiaProducto(tipologiaProducto);
                    }

                    #endregion

                    #region Guardar grados

                    foreach (GradoProductoInformacionBE grado in listaGrados.Where(x => x.CodProductoSAP.Equals(producto.idProducto)))
                    {
                        gradoPersistence.InsertGradoDef(grado);
                    }

                    #endregion

                    #region Guardar puntos

                    foreach (PuntosBE punto in listaPuntos.Where(x => x.idProducto.Equals(producto.idProducto)))
                    {
                        puntoPersistence.InsertPunto(punto);
                    }

                    #endregion

                    #region Guardar penalizacion regularidad producto

                    foreach (PenalizacionRegularidadProductoBE penalizacionProducto in listaPenalizacionesProducto.Where(x => x.idProducto.Equals(producto.idProducto)))
                    {
                        penalizacionProductoPersistence.InsertPenalizacionRegularidadProducto(penalizacionProducto);
                    }

                    #endregion

                    #region Guardar Destinos y tramos

                    foreach (DestinoBE destino in listaDestinos.Where(x => x.idProducto.Equals(producto.idProducto)))
                    {
                        destinoPersistence.InsertDestino(destino);

                        #region Guardar tramos

                        foreach (TramoBE tramo in listaTramos.Where(x => x.idDestino.Equals(destino.idDestino)))
                        {
                            tramoPersistence.InsertTramo(tramo);

                            #region Guardar descuentos de los tramos

                            foreach (DescuentoBE descuentoTramo in listaDescuentosTramos.Where(x => x.idTramo.Equals(tramo.idTramo)))
                            {
                                descuentoPersistence.InsertDescuentoTramo(descuentoTramo);
                            }

                            #endregion

                            #region Guardar tarifas de los tramos

                            foreach (TarifaBE tarifaProductoTramo in listaTarifasProductoTramo.Where(x => x.idTramo.Equals(tramo.idTramo)))
                            {
                                tarifaPersistence.InsertTarifaProducto(tarifaProductoTramo);
                            }

                            #endregion
                        }

                        #endregion

                        #region Guardar Descuentos Volumétrico                          

                        foreach (DescuentoBE descuentoVolumetrico in listaDescuentosVolumetrico.Where(x => x.idDestino.Equals(destino.idDestino)))
                        {
                            descuentoPersistence.InsertDescuentoVolumetrico(descuentoVolumetrico);
                        }

                        #endregion

                        #region Guardar Descuentos producto con destino

                        foreach (DescuentoBE descuentoProductoConDestino in listaDescuentosProductoConDestino.Where(x => x.idDestino.Equals(destino.idDestino)))
                        {
                            descuentoPersistence.InsertDescuentoProducto(descuentoProductoConDestino);
                        }

                        #endregion
                    }

                    #endregion                    

                    #region Guardar Descuentos producto sin destino

                    foreach (DescuentoBE descuentoProductoSinDestino in listaDescuentosProductoSinDestino.Where(x => x.idProducto.Equals(producto.idProducto)))
                    {
                        descuentoPersistence.InsertDescuentoProductoSinDestino(descuentoProductoSinDestino);
                    }

                    #endregion

                    #region Guardar tarifas producto sin destino

                    foreach (TarifaBE tarifaProductoSinDestinos in listaTarifasProductoSinDestino.Where(x => x.idProducto.Equals(producto.idProducto)))
                    {
                        tarifaPersistence.InsertTarifaProductoSinDestino(tarifaProductoSinDestinos);
                    }

                    #endregion

                    #region Valores añadidos

                    #region Guardar Relacion VA-Producto

                    foreach (ValorAnadidoAnexoProductoBE valorAnadidoProducto in listaValorAnyadidoProducto.Where(x => x.idProducto.Equals(producto.idProducto)))
                    {
                        valorAnadidoPersistence.InsertValorAnadidoProducto(valorAnadidoProducto);

                        #region Guardar VA

                        foreach (ValorAnadidoBE VA in listaValoresAnyadidos.Where(x => x.idValorAnadido.Equals(valorAnadidoProducto.idValorAnadido) && x.idValorAnadidoProducto.Equals(valorAnadidoProducto.idValorAnadidoAnexoProducto)))
                        {
                            valorAnadidoPersistence.InsertValorAnadido(VA);
                        }

                        #endregion

                        #region Guardar tipologias VA

                        foreach (TipologiaClienteBE tipologiaVA in listaTipologiasValorAnadido.Where(x => x.idValorAnadidoProducto.Equals(valorAnadidoProducto.idValorAnadidoAnexoProducto)))
                        {
                            tipologiaClientePersistence.InsertTipologiaVA(tipologiaVA);
                        }

                        #endregion

                        #region Guardar tarifas VA

                        foreach (TarifaBE tarifaVA in listaTarifasVA.Where(x => x.idValorAnadidoProducto.Equals(valorAnadidoProducto.idValorAnadidoAnexoProducto)))
                        {
                            tarifaPersistence.InsertTarifaVA(tarifaVA);
                        }

                        #endregion

                        #region Guardar Descuentos VA

                        foreach (DescuentoBE descuentoVA in listaDescuentosVA.Where(x => x.idValorAnadidoProducto.Equals(valorAnadidoProducto.idValorAnadidoAnexoProducto)))
                        {
                            descuentoPersistence.InsertDescuentoVA(descuentoVA);
                        }

                        // Recorremos la lista de características que tienen ListaValores que pertenecen a este valor añadido
                        foreach (CaracteristicaBE caracteristicaBE in listaCaracteristicasValoresAnadidos.Where(x => x.ListaValores.Any(y => y.idValorAnadido.Equals(valorAnadidoProducto.idValorAnadido))))
                        {
                            caracteristicasPersistence.InsertCaracteristicaValorAnadido(caracteristicaBE, (Guid)valorAnadidoProducto.idValorAnadido);
                        }

                        #endregion
                    }

                    #endregion

                    #endregion

                    #region Caracteristicas

                    #region Guardar relación caracteristicas - producto

                    foreach (ConfiguracionCaracteristicaBE caracteristicaProducto in listaCaracteristicasProducto.Where(x => x.idProducto.Equals(producto.idProducto)))
                    {
                        caracteristicasPersistence.InsertCaracteristicaProductoDeMemoria(caracteristicaProducto);
                    }

                    #endregion

                    #endregion

                    #region Guardar Rango población D2
                    RangoPoblacionD2BL rangoPoblacionD2BL = new RangoPoblacionD2BL();
                    
                    foreach (RangoPoblacionD2BE rangoPoblacionD2 in listaRangoPoblacionD2.Where(x => x.idProducto.Equals(producto.idProducto)))
                    {                        
                        rangoPoblacionD2BL.InsertRangosPoblacionD2(rangoPoblacionD2);
                    }

                    #endregion

                    //Se guarda en BD la definición del producto
                    uow.Save();

                }
            }
        }
       
        #endregion

        #region Métodos Eliminar

        public bool EliminarProductosNoSeleccionados(Guid idOferta, Collection<ProductoOfertaBE> productosSeleccionados, IUnitOfWork uow)
        {
            bool result = true;

            //obtengo todos los productos
            ProductoOfertaPersistence productoPersistencia = new ProductoOfertaPersistence(uow);
            Collection<ProductoOfertaBE> todosProductos = productoPersistencia.ObtenerProductosEnOferta(idOferta);

            Collection<ProductoOfertaBE> productosParaEliminar = new Collection<ProductoOfertaBE>();
            bool existe = false;


            //recorro todos los productos
            foreach (ProductoOfertaBE producto in todosProductos)
            {
                //recorro los productos seleccionados
                foreach (ProductoOfertaBE productoSeleccionado in productosSeleccionados)
                {
                    if (productoSeleccionado.idProducto == producto.idProducto && productoSeleccionado.ModalidadNegociacion == producto.ModalidadNegociacion)
                    {
                        existe = true;
                    }
                }
                //Si no existe en la lista de productos seleccionados lo añado a la lista de productos a eliminar.
                if (!existe)
                {
                    productosParaEliminar.Add(producto);
                }
            }

            //Recorro los productos a eliminar y comprueba que dentro de los seleccionados para eliminar no hay ninguno con un mismo Código de producto SAP
            //en la lista de productos seleccionados 

            //recorro todos los productos para eliminar
            foreach (ProductoOfertaBE productoEliminar in productosParaEliminar)
            {
                //recorro los productos seleccionados
                foreach (ProductoOfertaBE productoSeleccionado in productosSeleccionados)
                {
                    if (productoSeleccionado.CodProductoSAP == productoEliminar.CodProductoSAP)
                    {
                        productosParaEliminar.Remove(productoEliminar);
                    }
                }
            }

            //Eliminar Productos
            ProductoOfertaPersistence productoOfertaPersistencia = new ProductoOfertaPersistence(uow);
            productoOfertaPersistencia.EliminarProductos(productosParaEliminar);

            return result;
        }


        /// <summary>
        /// Metodo que evalua si lo que se quiere enviar es correcto
        /// </summary>
        /// <param name="productosSeleccionados">Lista de productos seleccionados</param>
        /// <returns>Retorna un objeto del tipo ResultadoCargaBE</returns>
        public ResultadoCargaBE ValidarDatos(Guid idOferta, Collection<ProductoOfertaBE> productosSeleccionados, string statusOferta, string codOferta)
        {
            ResultadoCargaBE objRespuesta = new ResultadoCargaBE();
            Collection<ErrorCargaBE> listaErrores = new Collection<ErrorCargaBE>();

            #region Validar si número de envíos es >0

            //Recorre los productos verificando si los productos tienen introducido un valor > 0
            foreach (ProductoOfertaBE productoSeleccionado in productosSeleccionados)
            {
                if (productoSeleccionado.NumeroEnvios <= 0)
                {
                    ErrorCargaBE errorEnvios = new ErrorCargaBE();
                    errorEnvios.producto = productoSeleccionado.CodProductoSAP;                    
                    errorEnvios.error = string.Format(CultureInfo.InvariantCulture, SimuladorResources.ErrorSincronizarFalloEnviosCero, productoSeleccionado.CodProductoSAP);
                    listaErrores.Add(errorEnvios); 
                }
            }

            #endregion

            #region validar numero de envios y devoluciones
            //Preparo las variables para la comprobación del numero de envíos y devoluciones
            decimal? numeroEnvios = 0;
            decimal? numeroDevoluciones = 0;

            //Preparo las variables para la comprobación de la existencia de al menos un producto devolucion en el caso de que se haya seleccionado alguno de enviso
            var relProdBL = new RelacionProductosBL();
            List<String> productosDevolucion = relProdBL.ObtenerProductosDevolucion();

            //Recorro los productos de devolución
            foreach (var codProductoDevolucion in productosDevolucion)
            {
                //Si el producto de devolución se encuentra entre los seleccionados
                ProductoOfertaBE productoSeleccionado = productosSeleccionados.FirstOrDefault(t => t.CodProductoSAP.Equals(codProductoDevolucion));

                if (productoSeleccionado != null)
                {
                    var listaProductosEnvioStr = relProdBL.ObtenerRelacionProductos(productoSeleccionado.CodProductoSAP, false);

                    //Obtenemos el listado de los productos de la oferta asociados a la devolución
                    var listaProductosEnvio = from prod in productosSeleccionados
                                              join prodEnvio in listaProductosEnvioStr
                                              on prod.CodProductoSAP equals prodEnvio
                                              select prod;

                    //Obtenemos los sumatorios de nº de envíos y de devoluciones
                    numeroEnvios = listaProductosEnvio.Sum(t => t.NumeroEnvios.HasValue ? t.NumeroEnvios.Value : 0);
                    numeroDevoluciones = productosSeleccionados.Where(t => t.CodProductoSAP.Equals(productoSeleccionado.CodProductoSAP))
                                                               .Sum(t => t.NumeroEnvios.HasValue ? t.NumeroEnvios.Value : 0);

                    //Si el número de envíos es menor que el de devoluciones mostramos un mensaje de error
                    if (numeroEnvios < numeroDevoluciones)
                    {
                        ErrorCargaBE errorDevoluciones = new ErrorCargaBE();
                        errorDevoluciones.error = string.Format(CultureInfo.InvariantCulture, SimuladorResources.ErrorDevolucionesEnvios + " Producto devolución: " + productoSeleccionado.CodProductoSAP  + ".");
                        listaErrores.Add(errorDevoluciones);
                    }
                }
            }
                
            //Collection<string> productos = new Collection<string>();
            //productos.Add(SimuladorResources.S0132);
            //productos.Add(SimuladorResources.S0133);
            //productos.Add(SimuladorResources.S0235);
            //productos.Add(SimuladorResources.S0236);

            //string devolucion = SimuladorResources.S0134;


            ////recorro los productos seleccionados y sumo los envios por un lado y las devoluciones por otro
            ////El numero de devoluciones no puede ser superios al de envios
            //foreach (ProductoOfertaBE productoSeleccionado in productosSeleccionados)
            //{
            //    if (productos.Contains(productoSeleccionado.CodProductoSAP))
            //    {
            //        numeroEnvios = numeroEnvios + productoSeleccionado.NumeroEnvios;
            //    }
            //    if (productoSeleccionado.CodProductoSAP == devolucion)
            //    {
            //        numeroDevoluciones = numeroDevoluciones + productoSeleccionado.NumeroEnvios;
            //    }
            //}


            //if (numeroEnvios < numeroDevoluciones)
            //{
            //    ErrorCargaBE errorDevoluciones = new ErrorCargaBE();
            //    errorDevoluciones.error = string.Format(CultureInfo.InvariantCulture, SimuladorResources.ErrorDevolucionesEnvios);
            //    listaErrores.Add(errorDevoluciones);                 
            //}

            ////Obtengo el numero de productos que conllevan devoluciones y obtengo tambien el numero de devoluciones
            //foreach (ProductoOfertaBE productoSeleccionado in productosSeleccionados)
            //{
            //    if (productos.Contains(productoSeleccionado.CodProductoSAP))
            //    {
            //        productosEnvios++;
            //    }
            //    if (productoSeleccionado.CodProductoSAP == devolucion)
            //    {
            //        productosDevoluciones++;
            //    }
            //}

            ////si hay algun producto de envíos y ninguno de deoluciones entonces error y viceversa...
            //if (productosEnvios > 0 && productosDevoluciones == 0)
            //{
            //    ErrorCargaBE errorDevoluciones = new ErrorCargaBE();
            //    errorDevoluciones.error = string.Format(CultureInfo.InvariantCulture, SimuladorResources.ErrorDevolucionesEnvios);
            //    listaErrores.Add(errorDevoluciones); 
            //}
            //else if (productosEnvios == 0 && productosDevoluciones > 0)
            //{
            //    ErrorCargaBE errorDevoluciones = new ErrorCargaBE();
            //    errorDevoluciones.error = string.Format(CultureInfo.InvariantCulture, SimuladorResources.ErrorDevolucionesEnvios);
            //    listaErrores.Add(errorDevoluciones); 
            //}
            #endregion

            #region validar si para todos los tramos de un productooferta si la modalidad negociacion es PRECIOcierto, si titne distribución tiene que tener la casilla rellena

            ProductoOfertaBL objProductoOfertaBL = new ProductoOfertaBL();
            ResultBE auxResult = objProductoOfertaBL.ComprobarPrecioCiertoTramosEnModalidad(productosSeleccionados);

            if (auxResult.Resultado == false)
            {
                ErrorCargaBE errorPreciosCiertosModalidad = new ErrorCargaBE();
                errorPreciosCiertosModalidad.error = auxResult.TextoError;
                listaErrores.Add(errorPreciosCiertosModalidad); 
            }

            #endregion

            #region validar si hay solapamiento de tramos para un productos y distintos Modelos de Negociación

            ValidacionesBL validacionesBL = new ValidacionesBL();
            auxResult = validacionesBL.ValidarSolapamientoTramosEnProducto(idOferta, productosSeleccionados);
            if (auxResult.Resultado == false)
            {
                ErrorCargaBE errorSolapamientoTramos = new ErrorCargaBE();
                errorSolapamientoTramos.error = auxResult.TextoError;
                listaErrores.Add(errorSolapamientoTramos); 
            }

            #endregion

            #region Valida si suma de Puntos y Grados vale 100%

            foreach (ProductoOfertaBE productoSeleccionado in productosSeleccionados)
            {
                auxResult = objProductoOfertaBL.ComprobarPuntosGradosPorcentajesListaProductos(productoSeleccionado);

                if (!auxResult.Resultado)
                {
                    ErrorCargaBE errorPuntosGradosPorcentajes = new ErrorCargaBE();
                    errorPuntosGradosPorcentajes.error = auxResult.TextoError;
                    listaErrores.Add(errorPuntosGradosPorcentajes);
                }
            }

            #endregion

            #region Validar regularidad en cero

            //recorro los productos seleccionados para validar si alguno tiene la regularidad en cero
            bool ProductoRegularidadCero = false;
            string codProductoRegularidadCero = string.Empty;

            foreach (ProductoOfertaBE productoSeleccionado in productosSeleccionados)
            {
                using (IUnitOfWork uow = new UnitOfWork())
                {
                    ConfiguracionGradoOfertaPersistence configuracionGradoOfertaPersistence = new ConfiguracionGradoOfertaPersistence(uow);
                    ConfiguracionGradoOfertaBE configuracionGradoOfertaBE = configuracionGradoOfertaPersistence.ObtenerConfiguracionGradoOferta(productoSeleccionado.idProductoOferta);
                    if (productoSeleccionado.Regularidad == 0 && configuracionGradoOfertaBE != null)
                    {
                        // TO DO: Mejorar esta chusta
                        if (!productoSeleccionado.CodProductoSAP.Equals("S0367"))
                        {
                            codProductoRegularidadCero = productoSeleccionado.CodProductoSAP;
                            ProductoRegularidadCero = true;
                            break;
                        }
                    }
                }

            }

            if (ProductoRegularidadCero)
            {
                ErrorCargaBE errorRegularidad = new ErrorCargaBE();
                errorRegularidad.error = String.Format(SimuladorResources.ErrorProductoRegularidadCero, codProductoRegularidadCero);
                listaErrores.Add(errorRegularidad);                

            }
            #endregion

            #region Validar relaciones productos

            ProductoOfertaBL objProductoOferta = new ProductoOfertaBL();

            bool estaVacio = string.IsNullOrWhiteSpace(codOferta) || codOferta.Equals("-");

            if ((!estaVacio && (statusOferta.Equals(SimuladorResources.CodigoEnProceso) || statusOferta.Equals(SimuladorResources.CodigoEnBorrador))) || estaVacio)
            {

                ResultBE respuestaComprobacion = objProductoOferta.ComprobarProductosSeleccionados(productosSeleccionados);

                if (!respuestaComprobacion.Resultado)
                {
                    ErrorCargaBE errorRelacion = new ErrorCargaBE();
                    errorRelacion.error = respuestaComprobacion.TextoError;
                    listaErrores.Add(errorRelacion);
                }
            }

            #endregion


            #region Validar que el descuento de cada destino está comprendido en [-999,999]
            
            auxResult = validacionesBL.ValidarDescuentosMaximosEnProducto(idOferta, productosSeleccionados);
            if (!auxResult.Resultado)
            {
                ErrorCargaBE errorRegularidad = new ErrorCargaBE();
                errorRegularidad.error = auxResult.TextoError;
                listaErrores.Add(errorRegularidad);                 
        }

        #endregion

            objRespuesta.errores = listaErrores;
            return objRespuesta;            
        }

        #endregion
    }
}
    