using Correos.SimuladorOfertas.Persistence;
using System.Data;

namespace Correos.SimuladorOfertas.Business
{
    public class GestorBaseDatosBL
    {
        /// <summary>
        /// Funcion que lanza la actualización del script de BBDD
        /// </summary>
        /// <param name="sqlConnectionString">cadena de conexión a la bbdd</param>
        /// <param name="sqlScriptFilePath">ruta donde se encuentra el script</param>
        public void ActualizarScriptBBDD(string sqlConnectionString, string sqlScriptFilePath)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                GestorBaseDatosPersistence GestorBBDDPersistence = new GestorBaseDatosPersistence(uow);
                GestorBBDDPersistence.ActualizarScriptBBDD(sqlConnectionString, sqlScriptFilePath);
            }
        }

        /// <summary>
        /// Funcion que lanza un script de BBDD
        /// </summary>
        /// <param name="sqlConnectionString">cadena de conexión a la bbdd</param>
        /// <param name="sqlScriptFilePath">script a ejecutar</param>
        public void EjecutaScriptBBDD(string sqlConnectionString, string sqlScriptFilePath)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                GestorBaseDatosPersistence GestorBBDDPersistence = new GestorBaseDatosPersistence(uow);
                GestorBBDDPersistence.EjecutaScriptBBDD(sqlConnectionString, sqlScriptFilePath);
            }
        }

        public bool EstaColumnaStatus(string sqlConnectionString, string sqlScriptFilePath)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                GestorBaseDatosPersistence GestorBBDDPersistence = new GestorBaseDatosPersistence(uow);
                return GestorBBDDPersistence.EstaColumnaStatus(sqlConnectionString, sqlScriptFilePath);
            }
        }

        /// <summary>
        /// Funcion que lanza un script de BBDD
        /// </summary>
        /// <param name="sqlConnectionString">cadena de conexión a la bbdd</param>
        /// <param name="sqlScriptFilePath">script a ejecutar</param>
        public bool ExisteColumaTablaBBDD(string sqlConnectionString, string tabla, string columna)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                GestorBaseDatosPersistence GestorBBDDPersistence = new GestorBaseDatosPersistence(uow);
                return GestorBBDDPersistence.ExisteColumaTablaBBDD(sqlConnectionString, tabla, columna);
            }
        }



        public DataTable select_sqlce(string sqlConnectionString, string sql)
        {
            using (IUnitOfWork uow = new UnitOfWork())
            {
                GestorBaseDatosPersistence GestorBBDDPersistence = new GestorBaseDatosPersistence(uow);
                return GestorBBDDPersistence.select_sqlce(sqlConnectionString, sql);
            }
        }



    }
}
