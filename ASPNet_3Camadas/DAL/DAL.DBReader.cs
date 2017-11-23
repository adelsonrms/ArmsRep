using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace DAL
{
    class DBReader: IFluentInterface
    {
        /// <summary>
        /// Executa um comando no banco de dados com uma conexão ativa e retorna um DataReader com os dados.
        /// </summary>
        /// <param name="sqlCommand">Comando SQL</param>
        /// <returns></returns>
        public IDataReader GetReader(string sqlCommand)
        {
            using (var cnn = DBConnection.Connection)
            {
                IDataReader red;
                try
                {
                    red = DBContext.GetData(sqlCommand, cnn);
                }
                //Erro especifico na abertura da conexão
                catch (DAL.Exceptions.DALExceptionConnectionOpen)
                {
                }
                //Erro especifico na conexçao
                catch (DAL.Exceptions.DALExceptionConnectionError)
                {
                }
                //Erro qualquer
                catch (Exception ex)
                {
                    throw new DAL.Exceptions.DALExceptionExecuteReader(ex.Message);
                }
                finally
                {
                    red=null;
                }
                //Retorno do reader
                return red;
            }
        }
    }
}
