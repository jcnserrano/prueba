using Correos.SimuladorOfertas.DTOs;
using Correos.SimuladorOfertas.Persistence;

namespace Correos.SimuladorOfertas.Business
{
    public class VersionesDataBaseBL
    {
        #region Metodos publicos

        /// <summary>
        /// Método que devuelve un registro con la información guardada en base de datos.
        /// </summary>
        /// <returns></returns>
        public VersionDataBaseBE ObtenerInformacionVersionBaseDatos()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                VersionesDataBasePersistence objPersistencia = new VersionesDataBasePersistence(uow);
                return objPersistencia.ObtenerInformacionVersionBBDD();
            }
        }

        /// <summary>
        /// desactiva el mostrado del mensaje de popup
        /// </summary>
        public void DesactivarMostrarMensajePopUp()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                VersionesDataBasePersistence objPersistencia = new VersionesDataBasePersistence(uow);
                objPersistencia.DesactivarMostrarMensajePopUp();
                uow.Save();
            }
        }

        /// <summary>
        /// Actualiza la fecha de actualizacion d definicion
        /// </summary>
        public void ActualizarFechaActualizarDefinicion()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                VersionesDataBasePersistence objPersistencia = new VersionesDataBasePersistence(uow);
                objPersistencia.ActualizarFechaActualizarDefinicion();
                uow.Save();
            }
        }

        #endregion


        public VersionDataBaseBE ObtenerDatosAuxiliares()
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                VersionesDataBasePersistence objPersistencia = new VersionesDataBasePersistence(uow);
                return objPersistencia.ObtenerDatosAuxiliares();
            }
        }

        public void actualizar_version_BDVersion(string sqlScriptFilePath)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                VersionesDataBasePersistence objPersistencia = new VersionesDataBasePersistence(uow);
                objPersistencia.actualizar_version_BDVersion(sqlScriptFilePath);

            }
        }
    }
}
