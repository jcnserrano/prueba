using Correos.SimuladorOfertas.Common.Extensions;
using Correos.SimuladorOfertas.DTOs;
using Correos.SimuladorOfertas.InOutLight;
using Correos.SimuladorOfertas.Persistence;
using System;
using System.Collections.ObjectModel;
using System.Linq;

namespace Correos.SimuladorOfertas.Business
{
    public class PotencialidadBL
    {
        #region Métodos Públicos

        /// <summary>
        /// Método que comprueba si para todos los productos y valores añadidos se tiene su potencialidad asociada
        /// </summary>
        /// <param name="usuario">Identificador del usuario</param>
        /// <param name="password">Contraseña del usuario</param>
        /// <param name="uow">Objeto transaccional</param>
        public void ComprobarPotencialidadesProductos(string usuario, string password, IUnitOfWork uow)
        {
            PotencialidadPersistence persistence = new PotencialidadPersistence(uow);
            Collection<PotencialidadBE> listaProductosVA = persistence.ObtenerProductosVASinPotencialidad(usuario);

            if (listaProductosVA != null && listaProductosVA.Count > 0)
            {
                if (SSOHelper.Instance.LogarConSSO)
                {
                    CommunicatorLight conectorSAP = new CommunicatorLight(SSOHelper.Instance.Usuario, password);
                    //SSOHelper.Instance.ActualizarCookiePortal();
                    SSOHelper.Instance.InicializarWSLight(conectorSAP.Cliente);
                    this.GuardarPotencialidades(conectorSAP.ZCObtenerPotencialidadRfc(listaProductosVA, usuario), usuario, uow);
                    SSOHelper.Instance.LimpiarWSLight();
                }
                else
                {
                    CommunicatorLight conectorSAP = new CommunicatorLight(usuario, password);
                    this.GuardarPotencialidades(conectorSAP.ZCObtenerPotencialidadRfc(listaProductosVA, usuario), usuario, uow);
                }
            }
        }

        /// <summary>
        /// Método que elimina la potencialidad de un producto
        /// </summary>
        /// <param name="idProducto">Identificador del producto</param>
        /// <param name="usuario">Código del usuario</param>
        public void EliminarPotencialidadProducto(Guid idProducto, string usuario, IUnitOfWork uow)
        {
            PotencialidadPersistence persistence = new PotencialidadPersistence(uow);
            persistence.EliminarPotencialidadProducto(idProducto, usuario);
        }

        /// <summary>
        /// Método que elimina la potencialidad de los valores añadidos de un producto y usuario
        /// </summary>
        /// <param name="idProducto">Identificador del producto</param>
        /// <param name="usuario">Código del usuario</param>
        public void EliminarPotencialidadVA(Guid idProducto, string usuario, IUnitOfWork uow)
        {
            PotencialidadPersistence persistence = new PotencialidadPersistence(uow);
            persistence.EliminarPotencialidadVA(idProducto, usuario);
        }

        #endregion

        #region Métodos Privados

        /// <summary>
        /// Método que guarda en base de datos toda la información de Potencialidades, tanto de productos como de valores añadidos
        /// </summary>
        /// <param name="listaPotencialidad">Lista de potencialidades</param>
        /// <param name="listaProductosBE">Lista de ProductoBE</param>
        /// <param name="usuario">Identificador del usuario</param>
        /// <param name="uow">Objeto transaccional</param>
        private void GuardarPotencialidades(Collection<PotencialidadBE> listaPotencialidad, string usuario, IUnitOfWork uow)
        {
            #region Productos

            // ------------------------
            // Se manejan los productos
            // ------------------------
            Collection<PotencialidadBE> listaPotencialidadesProductos = listaPotencialidad.Where(x => string.IsNullOrEmpty(x.VA)).ToList<PotencialidadBE>().ToCollection<PotencialidadBE>();

            foreach (PotencialidadBE potencialidadProducto in listaPotencialidadesProductos)
            {
                // Se gestiona el producto
                ProductoPersistence productoPersistence = new ProductoPersistence(uow);
                Collection<ProductoBE> productosDB = productoPersistence.ObtenerProductosByCodProductoSAP(potencialidadProducto.Producto);

                // El producto no puede ser nulo, no obstante se comprueba
                foreach (ProductoBE productoDB in productosDB)
                {
                    // Se gestiona la relación entre producto y usuario
                    PotencialidadPersistence potencialidadPersistence = new PotencialidadPersistence(uow);
                    PotencialidadBE potencialidadProductoDB = potencialidadPersistence.ObtenerPotencialidadProducto(productoDB.idProducto, usuario);

                    if (potencialidadProductoDB == null)
                    {
                        // Si no existe la relación, se crea
                        potencialidadProductoDB = new PotencialidadBE()
                        {
                            idPotencialidadProducto = Guid.NewGuid(),
                            Potencialidad = potencialidadProducto.Potencialidad,
                            idProducto = productoDB.idProducto,
                            Usuario = usuario
                        };
                    }

                    potencialidadPersistence.InsertUpdatePotencialidadProducto(potencialidadProductoDB);
                }
            }

            #endregion

            #region Valores Añadidos

            // -------------------------------
            // Se manejan los valores añadidos
            // -------------------------------
            Collection<PotencialidadBE> listaPotencialidadVA = listaPotencialidad.Except(listaPotencialidadesProductos).ToList<PotencialidadBE>().ToCollection<PotencialidadBE>();

            foreach (PotencialidadBE potencialidadVA in listaPotencialidadVA)
            {
                // Se gestiona el valor añadido
                ProductoPersistence productoPersistence = new ProductoPersistence(uow);

                Collection<ProductoBE> productosVADB = productoPersistence.ObtenerProductosByCodProductoSAP(potencialidadVA.Producto);

                // El producto no puede ser nulo, no obstante se comprueba 
                foreach (ProductoBE productoDB in productosVADB)
                {
                    ValorAnadidoProductoPersistence vapPersistence = new ValorAnadidoProductoPersistence(uow);
                    Guid idValorAnadidoProducto = vapPersistence.ObtenerIdValorAnadidoProductoByIdProductoCodVASAP(productoDB.idProducto, potencialidadVA.VA);

                    if (!idValorAnadidoProducto.Equals(Guid.Empty))
                    {
                        // Se gestiona la relación entre valor añadido y usuario
                        PotencialidadPersistence potencialidadPersistence = new PotencialidadPersistence(uow);
                        PotencialidadBE potencialidadVADB = potencialidadPersistence.ObtenerPotencialidadVA(idValorAnadidoProducto, usuario);

                        if (potencialidadVADB == null)
                        {
                            // Si no existe la relación, se crea
                            potencialidadVADB = new PotencialidadBE()
                            {
                                idPotencialidadVA = Guid.NewGuid(),
                                Potencialidad = potencialidadVA.Potencialidad,
                                idValorAnadidoProducto = idValorAnadidoProducto,
                                Usuario = usuario
                            };
                        }

                        potencialidadPersistence.InsertUpdatePotencialidadVA(potencialidadVADB);
                    }
                }
            }

            #endregion
        }

        #endregion
    }
}
