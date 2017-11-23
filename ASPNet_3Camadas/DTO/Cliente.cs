using Domain;
using System;
using System.Collections.Generic;

/// <summary>
/// Implementação das regras na Camada de Transferencia de Dados (Data Transfer Object)
/// </summary>
namespace DTO
{
    /// <summary>
    /// Define a implementação da entidade Cliente
    /// Herda da classe PropertiesController as funções encapsuladas para obter informações sobre os membros usando Reflections.
    /// Atributo Personalizado : Nessa classe é aplicado o recurso do atributo personalizado 'MappForField' que 
    /// </summary>
    public class Cliente : EntityController//, IDisposable
    {
        /// <summary>
        /// Invoca o método da classe base de controle de propriedades para que enumere todas as propriedades da classe filha em uma List
        /// tambem é possivel alterar/recuperar os valores das propriedades atraves do nome em forma de string.
        /// </summary>
        public Cliente(){base.SetChildClass(this);}
        /// <summary>
        /// Relaciona a lista de propriedades da entidade
        /// </summary>
        #region Propriedades
        public int ID { get; set; }
        public string Nome { get; set; }
        public string Endereco { get; set; }
        public string Telefone { get; set; }
        public string Email { get; set; }
        public string Observacoes { get; set; }
        #endregion
        
    }
}
