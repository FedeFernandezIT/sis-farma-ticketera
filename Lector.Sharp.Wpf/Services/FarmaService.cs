using System;
using System.IO;
using System.Linq;
using System.Reflection;
using Lector.Sharp.Wpf.Models;
using MySql.Data.MySqlClient;
using System.Configuration;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;

namespace Lector.Sharp.Wpf.Services
{
    public class FarmaService
    {
        public string Url { get; set; }
        public string Mostrador { get; set; }
        public string UrlNavegarCustom { get; set; }
        public string UrlMensajes { get; set; }
        public string UrlNavegar { get; set; }
        public string DatabaseServer { get; set; }
        public string DatabaseCatalog { get; set; }
        public string ShutdownServer { get; set; }
        public string ShutdownDatabase { get; set; }
        public string Presentation { get; set; }

        /// <summary>
        /// Lee los archivos de configuración y setea las propiedades correspondientes
        /// </summary>
        public void LeerFicherosConfiguracion()
        {
            try
            {
                var dir = ConfigurationManager.AppSettings["Directory.Setup"];

                var pathPresentation = Path.Combine(dir, ConfigurationManager.AppSettings["Presentation.Url"]);
                Presentation = new StreamReader(pathPresentation).ReadLine();

                var pathUrlInformacionRemoto = Path.Combine(dir, ConfigurationManager.AppSettings["Url.Informacion.Remoto"]);
                Url = new StreamReader(pathUrlInformacionRemoto).ReadLine();

                var pathUrlMensajesRemoto = Path.Combine(dir, ConfigurationManager.AppSettings["Url.Mensajes.Remoto"]);
                UrlMensajes = new StreamReader(pathUrlMensajesRemoto).ReadLine();

                var pathUrlCustom = Path.Combine(dir, ConfigurationManager.AppSettings["Url.Custom"]);
                UrlNavegarCustom = new StreamReader(pathUrlCustom).ReadLine();
                
                var pathDatabaseServer = Path.Combine(dir, ConfigurationManager.AppSettings["Database.Server"]);
                DatabaseServer = new StreamReader(pathDatabaseServer).ReadLine();

                var pathDatabseCatalog = Path.Combine(dir, ConfigurationManager.AppSettings["Database.Catalog"]);
                DatabaseCatalog = new StreamReader(pathDatabseCatalog).ReadLine();

                var pathShutdownServer = Path.Combine(dir, ConfigurationManager.AppSettings["Shutdown.Server"]);
                ShutdownServer = new StreamReader(pathShutdownServer).ReadLine();

                var pathShutdownDatabase = Path.Combine(dir, ConfigurationManager.AppSettings["Shutdown.Database"]);
                ShutdownDatabase = new StreamReader(pathShutdownDatabase).ReadLine();

                // Único archivo que puede no existir
                var pathMostradorVc = Path.Combine(dir, ConfigurationManager.AppSettings["Mostrador.Vc"]);                                
                Mostrador = File.Exists(pathMostradorVc) ? new StreamReader(pathMostradorVc).ReadLine() : "1";                

            } catch (IOException ex)
            {
                throw new IOException("Error al leer archivos de configuración");                
            }
        }

        /// <summary>
        /// Obtiene los códigos de barras de los medicametos y formatea la salida de los mismos.
        /// </summary>
        /// <returns>códigos de barras formateados</returns>
        public string[] GetCodigoBarraMedicamentos()

        {
            using (var db = new SisFarmaEntities(DatabaseServer, DatabaseCatalog))
            {

                var sql = "SELECT GROUP_CONCAT(DISTINCT SUBSTRING(cod_barras, 1, 3) SEPARATOR ';') AS codBarras FROM medicamentos WHERE NOT cod_barras IS NULL";
                var query = db.Database.SqlQuery<string>(sql).ToList();
                if (query.Count == 1)
                {
                    return query[0].Split(';');
                }
                return new string[0];                
            }
        }

        /// <summary>
        /// Obtiene los códigos de barras de los sinónimos y formatea la salida de los mismos.
        /// </summary>
        /// <returns>Códigos de barras formateados</returns>
        public string[] GetCodigoBarraSinonimos()
        {
            using (var db = new SisFarmaEntities(DatabaseServer, DatabaseCatalog))
            {
                var sql = "SELECT GROUP_CONCAT(DISTINCT SUBSTRING(cod_barras, 1, 3) SEPARATOR ';') AS codBarras FROM sinonimos WHERE NOT cod_barras IS NULL";
                var query = db.Database.SqlQuery<string>(sql).ToList();
                if (query.Count == 1)
                {
                    return query[0].Split(';');
                }
                return new string[0];
            }
        }

        /// <summary>
        /// Obtiene el código de un cliente con una determinada tarjeta
        /// </summary>
        /// <param name="tarjeta"></param>
        /// <returns></returns>
        public long? GetCliente(string tarjeta)
        {
            using (var db = new SisFarmaEntities(DatabaseServer, DatabaseCatalog))
            {
                var sql = "SELECT cod FROM clientes WHERE tarjeta = '" + tarjeta + "'";
                var query = db.Database.SqlQuery<long>(sql).ToList();
                if (query.Count != 0)
                {
                    return query[0];
                }
                return null;
            }
        }

        /// <summary>
        /// Obtiene el identificador de un trabajador con una determinada tarjeta
        /// </summary>
        /// <param name="tarjeta"></param>
        /// <returns></returns>
        public long? GetTrabajador(string tarjeta)
        {
            using (var db = new SisFarmaEntities(DatabaseServer, DatabaseCatalog))
            {
                var sql = "SELECT id FROM trabajador WHERE tarjeta = '" + tarjeta + "'";
                var query = db.Database.SqlQuery<long>(sql).ToList();
                if (query.Count != 0)
                {
                    return query[0];
                }
                return null;
            }
        }

        /// <summary>
        /// Obtiene el código nacional de un sinónimo según un filtro aplicado a los códigos de barra
        /// </summary>
        /// <param name="filter"></param>
        /// <returns>código nacioanal</returns>
        public long? GetCodigoNacionalSinonimo(string filter)
        {
            using (var db = new SisFarmaEntities(DatabaseServer, DatabaseCatalog))
            {
                var sql = "SELECT cod_nacional FROM sinonimos WHERE SUBSTRING(cod_barras, 1, 12) LIKE '" + filter + "%'";
                var query = db.Database.SqlQuery<string>(sql).ToList();
                if (query.Count != 0)
                {
                    return Convert.ToInt64(query[0]);
                }
                return null;                
            }            
        }

        /// <summary>
        /// Obtiene el código nacional de un medicamento según un filtro aplicado a los códigos de barra
        /// </summary>
        /// <param name="filter"></param>
        /// <returns>código nacioanal</returns>
        public long? GetCodigoNacionalMedicamento(string filter)
        {
            using (var db = new SisFarmaEntities(DatabaseServer, DatabaseCatalog))
            {
                var sql = "SELECT cod_nacional FROM medicamentos WHERE SUBSTRING(cod_barras, 1, 12) LIKE '" + filter + "%'";
                var query = db.Database.SqlQuery<long>(sql).ToList();                
                if (query.Count != 0)
                {
                    return query[0];
                }
                return null;
            }
        }

        /// <summary>
        /// Obtiene un asociado en la ventas por el código nacional.
        /// </summary>
        /// <param name="codNacional"></param>
        /// <returns>Identificador de Asociado</returns>
        public string GetAsociado(long codNacional)
        {
            using (var db = new SisFarmaEntities(DatabaseServer, DatabaseCatalog))
            {
                var sql = "SELECT a.asociado FROM ventas_cruzadas v " +
                            "INNER JOIN asociados_cruzadas a ON a.idVentaCruzada = v.id " +
                            "WHERE a.asociado = '" + codNacional + "' " +
                            "AND v.eliminado = 0 AND v.activo = 1 LIMIT 0,1";
                var query = db.Database.SqlQuery<string>(sql).ToList();                    
                return query.Count != 0 ? query.First() : null;                               
            }
        }

        /// <summary>
        /// Obtiene un artículo de las ventas según el código nacional
        /// </summary>
        /// <param name="codNacional"></param>
        /// <returns></returns>
        public string GetArticulo(long codNacional)
        {
            using (var db = new SisFarmaEntities(DatabaseServer, DatabaseCatalog))
            {
                var sql = "SELECT la.cod_articulo AS asociado " +
                            "FROM ventas_cruzadas v " +
                            "INNER JOIN asociados_cruzadas a ON a.idVentaCruzada = v.id " +
                            "INNER JOIN listas_articulos la ON la.cod_lista = a.asociado " +
                            "WHERE la.cod_articulo = '" + codNacional + "' AND v.tipoAsociado = 'Por Listas' " +
                            "AND v.eliminado = 0 AND v.activo = 1 LIMIT 0,1";
                var query = db.Database.SqlQuery<int>(sql).ToList();
                return query.Count != 0 ? query[0].ToString() : null;                
            }
        }

        /// <summary>
        /// Obtiene el cdódigo nacional de la primera categorización
        /// </summary>
        /// <returns></returns>
        public long? GetCategorizacion()
        {
            using (var db = new SisFarmaEntities(DatabaseServer, DatabaseCatalog))
            {
                var sql = "SELECT cod_nacional FROM categorizacion LIMIT 0,1";
                var query = db.Database.SqlQuery<long>(sql).ToList();
                if (query.Count != 0)
                {
                    return query[0];
                }
                return null;
            }
        }

        /// <summary>
        /// Obtiene el código nacional de un asociado categorizado con ventas, según un código nacioanal
        /// </summary>
        /// <param name="codNacional"></param>
        /// <returns></returns>
        public string GetAsociadoCategorizacion(long codNacional)
        {
            using (var db = new SisFarmaEntities(DatabaseServer, DatabaseCatalog))
            {
                var sql = "SELECT c.cod_nacional AS asociado " +
                            "FROM ventas_cruzadas v " +
                                "INNER JOIN asociados_cruzadas a ON a.idVentaCruzada = v.id " +
                                "INNER JOIN categorizacion c ON IF(INSTR(a.asociado,':') > 0, c.subfamilia = SUBSTRING_INDEX(a.asociado,':',-1), 1 = 1) AND c.familia = SUBSTRING_INDEX(a.asociado,':',1) " +
                                      "WHERE c.cod_nacional = @codNacional AND v.tipoAsociado = 'Por Familia/Subfamilia' " +
                                          "AND v.eliminado = 0 AND v.activo = 1 LIMIT 0,1";
                var query = db.Database.SqlQuery<long>(sql, new MySqlParameter("@codNacional", codNacional)).ToArray();
                return query.Length != 0 ? query[0].ToString() : null;                                
            }
        }

        /// <summary>
        /// Obtiene el código nacional de un asociado con ventas de medicamentos, según un código nacional
        /// </summary>
        /// <param name="codNacional"></param>
        /// <returns>Código nacional</returns>
        public string GetAnyAsociadoMedicamento(long codNacional)
        {
            using (var db = new SisFarmaEntities(DatabaseServer, DatabaseCatalog))
            {
                var sql = "SELECT m.cod_nacional AS asociado " +
                            "FROM ventas_cruzadas v " +
                                "INNER JOIN asociados_cruzadas a ON a.idVentaCruzada = v.id " +
                                "INNER JOIN medicamentos m ON m.familia = SUBSTRING_INDEX(a.asociado,':',-1) OR m.laboratorio = a.asociado " +
                                "WHERE m.cod_nacional = @codNacional AND (v.tipoAsociado = 'Por Familia/Subfamilia' OR v.tipoAsociado = 'Por Laboratorio') " +
                                        "AND v.eliminado = 0 AND v.activo = 1 LIMIT 0,1";
                var query = db.Database.SqlQuery<long>(sql, new MySqlParameter("@codNacional", codNacional)).ToArray();
                return query.Length != 0 ? query[0].ToString() : null;
            }
        }

        public void CheckShutdownTime()
        {
            var now = DateTime.Now;
            var horarios = GetTimesShutdown();
            foreach (var horario in horarios)
            {                
                if (horario.Hora == now.Hour && horario.Minutos == now.Minute)
                {
                    Process.Start("shutdown", "/s /t 30");
                    break;
                }                    
            }
        }

        private List<HoraMinutos> GetTimesShutdown()
        {
            try
            {
                using (var db = new ShutdownContext(ShutdownServer, ShutdownDatabase))
                {
                    var sql = @"select hora as Hora, minutos as Minutos from hora_apagado";
                    return db.Database.SqlQuery<HoraMinutos>(sql).ToList();
                }
            }
            catch (Exception)
            {
                return new List<HoraMinutos>();
            }            
        }
    
    }
}