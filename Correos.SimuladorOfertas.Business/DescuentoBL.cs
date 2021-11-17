using Correos.SimuladorOfertas.DTOs;
using Correos.SimuladorOfertas.Persistence;
using System;
using System.Collections.ObjectModel;
using System.Collections.Generic;

namespace Correos.SimuladorOfertas.Business
{
    public class DescuentoBL
    {
        #region Métodos privados

        /// <summary>
        /// Método que obtiene el descuento correspondiente al destino segun modelo de costes, publicorreo, libros y publicaciones
        /// </summary>
        /// <param name="idDestino">Identificador del destino</param>
        /// <param name="idTipoCliente">Identificador del tipo de cliente</param>
        /// <returns>Descuento máximo para el destino y tipo de cliente</returns>
        private double ObtenerDescuentoDestino(Guid idDestino, ConfiguracionGradoOfertaBE objConfGradoOferta, ConfiguracionPuntoOfertaBE objConfPuntoOferta, int regularidadUsuario, out double descuentoMaxDestino, string tipoCliente)
        {
            double objDescuento = 0;

            using (IUnitOfWork uow = new UnitOfWork())
            {
                DescuentoPersistence descuentopersistence = new DescuentoPersistence(uow);
                descuentoMaxDestino = descuentopersistence.ObtenerDescuentoDestino(idDestino, tipoCliente);

                if (objConfGradoOferta != null)
                {
                    PuntoPersistence puntopersistence = new PuntoPersistence(uow);
                    PuntosBE objPuntos = puntopersistence.ObtenerPuntoAsociadoAlDestino(idDestino);

                    GradoPersistence gradopersistence = new GradoPersistence(uow);
                    GradosBE objGrados = gradopersistence.ObtenerGradoAsociadoAlDestino(idDestino);

                    ProductoPersistence productopersistence = new ProductoPersistence(uow);

                    decimal regularidadUsuarioDecimal = 0;
                    decimal.TryParse(regularidadUsuario.ToString(), out regularidadUsuarioDecimal);


                    decimal? regularidad = productopersistence.ObtenerPenalizacionProducto(idDestino, regularidadUsuarioDecimal);

                    if (regularidad == null)
                    {
                        int regularidadproducto = productopersistence.ObtenerRegularidadProducto(idDestino);
                        decimal constante = 34.61m;

                        decimal regularidadproductoDecimal = 0;
                        decimal.TryParse(regularidadproducto.ToString(), out regularidadproductoDecimal);

                        decimal division = regularidadUsuarioDecimal / regularidadproductoDecimal;
                        decimal aux = 1 - division;
                        regularidad = constante * aux;

                        if (regularidad < 0)
                        {
                            regularidad = 0;
                        }

                    }

                    decimal? penalizaciones = objConfGradoOferta.DistribucionG0 * objGrados.G0 + objConfGradoOferta.DistribucionG1 * objGrados.G1 + objConfGradoOferta.DistribucionG2 * objGrados.G2 +
                        objConfPuntoOferta.DistribucionCAM * objPuntos.CAM + objConfPuntoOferta.DistribucionRUR * objPuntos.RUR + objConfPuntoOferta.DistribucionURB * objPuntos.URB + regularidad.Value;

                    double penalizacion;
                    double.TryParse(penalizaciones.ToString(), out penalizacion);

                    objDescuento = descuentoMaxDestino - penalizacion;
                    if (objDescuento < 0)
                    {
                        objDescuento = 0;
                    }
                }
                else
                {
                    return 0;
                }
            }

            return objDescuento;
        }
        #endregion

        #region Método Obtener

        /// <summary>
        /// obtiene el descuento máximo de un producto sin destinos
        /// </summary>
        /// <param name="idProducto"></param>
        /// <param name="tipoCliente"></param>
        /// <returns></returns>
        public double ObtenerDescuentoSinDestino(Guid idProducto, string tipoCliente)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                DescuentoPersistence persistence = new DescuentoPersistence(uow);
                return persistence.ObtenerDescuentoSinDestino(idProducto, tipoCliente);
            }
        }

        /// <summary>
        /// Obtiene la lista de descuentos de los tramos de un producto dependiendo del tiop de cliente
        /// </summary>
        /// <param name="tipoCliente"></param>
        /// <param name="idProducto"></param>
        /// <returns></returns>
        public Collection<DescuentoTramoBE> ObtenerDescuentoTramos(string tipoCliente, Guid idProducto)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                DescuentoPersistence persistence = new DescuentoPersistence(uow);
                return persistence.ObtenerTramoDescuento(idProducto, tipoCliente);
            }
        }

        /// <summary>
        /// Método que obtiene el descuento correspondiente al destino en función del tipo del cliente
        /// </summary>
        /// <param name="idDestino">Identificador del destino</param>
        /// <param name="idTipoCliente">Identificador del tipo de cliente</param>
        /// <returns>Descuento máximo para el destino y tipo de cliente</returns>
        public double ObtenerDescuentoDestino(Guid idDestino, string tipoCliente)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                DescuentoPersistence persistence = new DescuentoPersistence(uow);
                return persistence.ObtenerDescuentoDestino(idDestino, tipoCliente);
            }
        }

        //Obtiene la regularidad de un producto
        public Collection<RegularidadVolumetricoBE> ObtenerBonificacionRegularidad(Guid idProducto)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                ProductoPersistence productopersistence = new ProductoPersistence(uow);
                return productopersistence.ObtenerBonificacionesProducto(idProducto);
            }
        }

        /// <summary>
        /// Método que obtiene el descuento correspondiente al destino segun modelo de costes
        /// </summary>
        /// <param name="idDestino">Identificador del destino</param>
        /// <param name="idTipoCliente">Identificador del tipo de cliente</param>
        /// <returns>Descuento máximo para el destino y tipo de cliente</returns>
        public double ObtenerDescuentoDestinoSegunCostes(Guid idDestino, ConfiguracionGradoOfertaBE objConfGradoOferta, ConfiguracionPuntoOfertaBE objConfPuntoOferta, int regularidadUsuario, out double descuentoMaxDestino, string tipoCliente)
        {
            return this.ObtenerDescuentoDestino(idDestino, objConfGradoOferta, objConfPuntoOferta, regularidadUsuario, out descuentoMaxDestino, tipoCliente);
        }

        /// <summary>
        /// Método que obtiene el descuento correspondiente al destino segun modelo de libros
        /// </summary>
        /// <param name="idDestino">Identificador del destino</param>
        /// <param name="idTipoCliente">Identificador del tipo de cliente</param>
        /// <returns>Descuento máximo para el destino y tipo de cliente</returns>
        public double ObtenerDescuentoDestinoSegunLibros(Guid idDestino, ConfiguracionGradoOfertaBE objConfGradoOferta, ConfiguracionPuntoOfertaBE objConfPuntoOferta, int regularidadUsuario, out double descuentoMaxDestino, string tipoCliente)
        {
            return this.ObtenerDescuentoDestino(idDestino, objConfGradoOferta, objConfPuntoOferta, regularidadUsuario, out descuentoMaxDestino, tipoCliente);
        }

        /// <summary>
        /// Método que obtiene el descuento correspondiente al destino segun modelo de publicaciones
        /// </summary>
        /// <param name="idDestino">Identificador del destino</param>
        /// <param name="idTipoCliente">Identificador del tipo de cliente</param>
        /// <returns>Descuento máximo para el destino y tipo de cliente</returns>
        public double ObtenerDescuentoDestinoSegunPublicaciones(Guid idDestino, ConfiguracionGradoOfertaBE objConfGradoOferta, ConfiguracionPuntoOfertaBE objConfPuntoOferta, int regularidadUsuario, out double descuentoMaxDestino, string tipoCliente)
        {
            return this.ObtenerDescuentoDestino(idDestino, objConfGradoOferta, objConfPuntoOferta, regularidadUsuario, out descuentoMaxDestino, tipoCliente);
        }

        /// <summary>
        /// Método que obtiene el descuento correspondiente al destino segun modelo de publicorreo
        /// </summary>
        /// <param name="idDestino">Identificador del destino</param>
        /// <param name="idTipoCliente">Identificador del tipo de cliente</param>
        /// <returns>Descuento máximo para el destino y tipo de cliente</returns>
        public double ObtenerDescuentoDestinoSegunPublicorreo(Guid idDestino, ConfiguracionGradoOfertaBE objConfGradoOferta, ConfiguracionPuntoOfertaBE objConfPuntoOferta, int regularidadUsuario, out double descuentoMaxDestino, string tipoCliente)
        {
            return this.ObtenerDescuentoDestino(idDestino, objConfGradoOferta, objConfPuntoOferta, regularidadUsuario, out descuentoMaxDestino, tipoCliente);
        }

        /// <summary>
        /// Método que obtiene el descuento correspondiente al valor añadido en función del tipo del cliente
        /// </summary>
        /// <param name="idValorAnadidoProducto">Identificador del valor añadido producto</param>
        /// <param name="idTipoCliente">Identificador del tipo de cliente</param>
        /// <returns>Descuento máximo para el valor añadido y tipo de cliente</returns>
        public double ObtenerDescuentoVA(Guid idValorAnadidoProducto, string tipoCliente)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                DescuentoPersistence persistence = new DescuentoPersistence(uow);
                return persistence.ObtenerDescuentoVA(idValorAnadidoProducto, tipoCliente);
            }
        }

        #endregion


        /// <summary>
        /// Metodo que devuelve la lista de descuentos máximos de los destinos asociados al producto
        /// </summary>
        /// <param name="anexoSAP"></param>
        /// <param name="productoSAP"></param>
        /// <returns></returns>
        public Collection<DescuentoMaximoVolumetricoBE> ObtenerDescuentosProducto(string anexoSAP, string productoSAP)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                DescuentoPersistence persistence = new DescuentoPersistence(uow);
                return persistence.ObtenerDescuentosProducto(anexoSAP, productoSAP);
            }

        }

        /// <summary>
        /// Método que obtiene todos los descuentos de la tabla DescuentoTramo para un destino
        /// </summary>
        /// <param name="idDestino"></param>
        /// <returns></returns>
        public Collection<DescuentoBE> ObtenerDescuentosTramosByIdDestino(Guid idDestino)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                return ObtenerDescuentosTramosByIdDestino(idDestino, uow);
            }
        }

        /// <summary>
        /// Método que obtiene todos los descuentos de la tabla DescuentoTramo para un destino
        /// </summary>
        /// <param name="idDestino"></param>
        /// <returns></returns>
        public Collection<DescuentoBE> ObtenerDescuentosTramosByIdDestino(Guid idDestino, IUnitOfWork uow)
        {
            DescuentoPersistence persistence = new DescuentoPersistence(uow);
            return persistence.ObtenerDescuentosTramosByIdDestino(idDestino);
        }

        /// <summary>
        /// Método que obtiene los descuentos de la tabla DescuentoVolumetrico de un destino
        /// </summary>
        /// <param name="idDestino">Identificador del destino</param>
        /// <returns>Colección de entidades DescuentoBE</returns>
        public Collection<DescuentoBE> ObtenerDescuentosVolumetricoByIdDestino(Guid idDestino)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                return ObtenerDescuentosVolumetricoByIdDestino(idDestino, uow);
            }
        }

        /// <summary>
        /// Método que obtiene los descuentos de la tabla DescuentoVolumetrico de un destino
        /// </summary>
        /// <param name="idDestino">Identificador del destino</param>
        /// <returns>Colección de entidades DescuentoBE</returns>
        public Collection<DescuentoBE> ObtenerDescuentosVolumetricoByIdDestino(Guid idDestino, IUnitOfWork uow)
        {
            DescuentoPersistence persistence = new DescuentoPersistence(uow);
            return persistence.ObtenerDescuentosVolumetricoByIdDestino(idDestino);
        }

        /// <summary>
        /// Obtener todos los descuentos para la lista de destinos
        /// </summary>
        /// <param name="listaIdsProductos"></param>
        /// <returns></returns>
        public Dictionary<Guid, Collection<DescuentoBE>> ObtenerDescuentosDestinos(Collection<Guid> listaIdsDestinos)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                DescuentoPersistence destinoPersistence = new DescuentoPersistence(uow);
                return destinoPersistence.ObtenerDescuentosProductos(listaIdsDestinos);
            }
        }
    }
}
