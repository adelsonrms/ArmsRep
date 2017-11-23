using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using DTO;

namespace BLL
{
   public class ClienteBLL : ICliente<Cliente>
    {

        #region Operações CRUD
            public void Alterar(Cliente obj)
            {
                throw new NotImplementedException();
            }
            public void Excluir(Cliente obj)
            {
                throw new NotImplementedException();
            }
            public void Incluir(Cliente obj)
            {
                throw new NotImplementedException();
            }
            public void Inserir(Cliente obj)
            {
                throw new NotImplementedException();
            }
        #endregion

        #region Consultas
        //DONE: Falta implementar : ConsultarPorNome()
        public DataTable ConsultarPorNome(string nome)
            {
            var dt = new DataTable();
            dt = DAL.DBContext.GetDataTable(new Cliente().TSQLSelectByField(fieldName:"Nome", fieldValue:nome));
            return dt;
        }

        public DataTable ConsultaPorID(int id)
        {
            var dt = new DataTable();
            try
            {
                dt = DAL.DBContext.GetDataTable(new Cliente().TSQLSelectByID(id: id));
            }
            catch
            {

            }
            return dt;
        }

        public DataTable ConsultaPorNome(string nome)
        {
            var dt = new DataTable();
            try
            {
                dt = DAL.DBContext.GetDataTable(new Cliente().TSQLSelectByField(fieldName: "nome", fieldValue:nome));
            }
            catch
            {

            }
            return dt;
        }
        /// <summary>
        /// Retorna a instancia de um Cliente passando o ID
        /// </summary>
        /// <param name="id">ID do Cliente</param>
        /// <returns></returns>
        public Cliente GetClienteByID(int id)
            {
                try
                {
                    var cliente = new Cliente();
                    var dt = DAL.DBContext.GetDataTable(cliente.TSQLSelectByID(id: id));
                    return populaCliente(cliente, dt);
                }
                catch (Exception ex)
                {
                    return null;
                    throw new Exception("Erro ao obter o Cliente : " + ex.Message);
                }
            }
        #endregion

        #region Listas

        //DONE: Falta implementar : ExibirTodos()
            public DataTable ExibirTodos()
            {
                var dt = new DataTable();
                dt = DAL.DBContext.GetDataTable(new Cliente().TSQLSelectByField());
                return dt;
            } 

            public IList<Cliente> Exibir()
            {
                var cliente = new Cliente();
                DataTable dt = DAL.DBContext.GetDataTable(cliente.TSQLSelectByField());
                return GetClientesFromDataTable(dt);
            }

            private IList<Cliente> GetClientesFromDataTable(DataTable fromTable)
            {
                var count = fromTable.Rows.Count;
                var lstClientes = new List<Cliente>();

                if (count >0 )
                {
                    foreach (var linha in fromTable.Rows)
                    {
                        var cliente = populaCliente(new Cliente(), fromTable);
                        if (cliente!=null)
                        {
                            lstClientes.Add(cliente);
                        }
                    }
                }
                return lstClientes;
            }
        #endregion

        #region Suporte
        private Cliente populaCliente(Cliente cliente  , DataTable fromTable)
        {
            try
            {
                //using (
                var cli = cliente;
                    //)
                
                    var dataC = fromTable.Rows[0].ItemArray;

                    for (int coluna = 0;
                        coluna < fromTable.Columns.Count - 1;
                        coluna++)
                    {
                        var valor = dataC[fromTable.Columns[coluna].Ordinal];

                        if (!valor.Equals(DBNull.Value))
                        {
                            cli.setValue(fromTable.Columns[coluna].ColumnName, valor);
                        }
                    }
                    return cli;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        #endregion

    }
}
