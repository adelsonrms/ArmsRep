//Namespaces de acesso a dados
//Usar somente o necessário
using System;
using System.Data;
using System.Data.SqlClient;
using System.Runtime.Serialization;
using DTO;
using System.Configuration;
using DAL.Exceptions;
using System.ComponentModel;
/// <summary>
/// Namespace onde ficam as classes da camada de acesso a dados
/// </summary>
namespace DAL
{


    public enum ConnectionState
    {
        Opened, Close
    }
    /// <summary>
    /// Implementa as funcionalidades que interage com o banco de dados.
    /// </summary>
    public class DBConnection:IDisposable
    {
        static public SqlConnection Connect()
        {return MakeConnect();
        }

        /// <summary>
        /// Variavel local privada que manterá a instancia da Conexão
        /// </summary>
        static private SqlConnection _cnn;
        /// <summary>
        /// Recupera o status da conexão (Somente leitura)
        /// </summary>
        public ConnectionState ConnectionState {
            get
            {
                if (_cnn!=null)
                {
                    return (ConnectionState)_cnn.State;
                }
                else
                    {
                    return ConnectionState.Close;
                };
            }
        }

        /// <summary>
        /// Retorna a conexão aberta. Caso nao esteja aberta, abre-a
        /// </summary>
        static public SqlConnection Connection
        {
            get {
                if (_cnn == null) {MakeConnect();}
                return _cnn;
            }
        }
        /// <summary>
        /// Efetua a conexão com o banco de dados atraves da string de conexão armazenada no Web.Config
        /// </summary>
        /// <returns>Retorna um objeto SqlConnection aberto (ou não em caso de falha)</returns>
        static private SqlConnection MakeConnect()
        {
            try
            {
                //Inicializa uma nova conexão
                _cnn = new SqlConnection(ConfigurationManager.ConnectionStrings["cS_Cliente_LocalDB"].ConnectionString);
                _cnn.Open();
            }
            catch (Exception sex)
            {
                _cnn = null;
                throw new DALExceptionConnectionOpen(string.Format("Ocorreu um erro ao realizar a conexão com o banco de dados.\n Detalhes : {0}", sex.Message));
            }
            return _cnn;
        }

        /// <summary>
        /// Inicializa uma conexão permitindo forçar a criação de uma nova instancia
        /// </summary>
        /// <param name="bForceNew">Se true, a conexão atual será finalizada e uma nova será aberta</param>
        /// <returns></returns>
        private SqlConnection GetConnection(bool bForceNew = false)
        {
            if (bForceNew)
            {
                CloseConnection();
            };
            return MakeConnect();
        }

        /// <summary>
        /// Finaliza a conexão atualmente ativa
        /// </summary>
        /// <returns></returns>
        private bool CloseConnection()
        {
            var ret = true;
            try
            {
                if (_cnn != null)
                {
                    _cnn.Close();
                }
            }
            catch (Exception ex)
            {
                ret = false;
                throw new DAL.Exceptions.DALExceptionConnectionError(string.Format("Ocorreu um erro ao finalizar a conexão com o banco de dados\n Erro : {0}", ex.Message));
            }
            finally
            {
                _cnn = null;
                _cnn.Dispose();
            }
            return ret;
        }

        public void Dispose()
        {
            this.Dispose();
        }
    }
}
