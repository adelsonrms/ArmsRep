using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DTO; //Referencia onde esta a classe Cliente
using System.Data;

namespace BLL
{
    /// <summary>
    /// Interface que contem os contratos (nome dos metodos) das regras de negocios
    /// <T> indica que a interface (ou classe) será representada como  uma coleção de obejtos (Generic List)
    /// Onde : T Representará o objeto que utilizará a lista. T é a representação padrão. pderia ser qualquer letra
    /// Nesse casso, caso T será uma classe Cliente
    /// </summary>
    interface ICliente<T>
    {
        DataTable ExibirTodos(); //Metodo que retornará um DataTable
        IList<T> Exibir(); //metodo que retornará uma lista inteira de objetos (Cliente)

        void Incluir(T obj);
        void Alterar(T obj);
        void Excluir(T obj);
        void Inserir(T obj);

        DataTable ConsultarPorNome(string nome);
        Cliente GetClienteByID(int id);
    }
}
