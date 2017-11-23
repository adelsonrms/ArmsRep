using System;
using System.Collections.Generic;
using System.Linq;

namespace DTO
{
    /// <summary>
    /// Representa a classe que deverá ser herdada por outras para obter informações automaticas das propriedades
    /// </summary>
    public class EntityController
    {
        private object thisClass;
        private List<Property> _Properties;

        public List<Property> Properties {
            get {
                if (_Properties == null)
                {
                    _Properties = GetProperties(thisClass);
                }
                return _Properties
          ; }
            set { _Properties = value; }
        }
        List<string> propString { get; set; }

        /// <summary>
        /// Construtor base para inicialização do controler
        /// </summary>
        public EntityController()
        {
        }
        /// <summary>
        /// Ao passar uma instancia de uma Classe, recupera a lista de propriedades da classe
        /// </summary>
        /// <param name="objClass"></param>
        public EntityController(object objClass)
        {
            GetProperties(objClass);
        }
        /// <summary>
        /// Define a instancia da classe a ser analisada
        /// </summary>
        /// <param name="objClass">Instancia da classe</param>
        internal void SetChildClass(object objClass)
        {
            thisClass = objClass;
        }
        
        #region "Consultas SELECT abstraidas"

        //DONE : Feito
        /// <summary>
        /// Recupera a String do SELECT ALL FROM TABLE referente à classe
        /// </summary>
        private string TSQLSelectAll
        {
            get{
                return string.Format("SELECT {0} FROM {1}", this.GetSelectFields(), thisClass.GetType().Name);
            }
        }

        public string TSQLInsertAll
        {
            get
            {
                return string.Format("INSERT INTO {0} ({1}) VALUES ({2})", thisClass.GetType().Name, this.GetSelectFields(), this.GetInsertFields());
            }
        }


        /// <summary>
        /// Monta uma String SELECT considerando todos os campos (SELECT * FROM) e ja filtra por ID especificando o valor do ID
        /// </summary>
        /// <param name="id">ID a ser filtrado</param>
        /// <returns>String com o SELECT montado</returns>
        public string TSQLSelectByID(int id) 
        {
           return string.Format("{0} WHERE ID={1}", TSQLSelectAll, id);
        }
        /// <summary>
        /// Monta a string SELECT ja assumindo o ID atual da classe
        /// </summary>
        /// <returns>String com o SELECT montado</returns>
        public string TSQLSelectByID()
        {
            return TSQLSelectByID(int.Parse(getValue("ID").ToString()));
        }
        /// <summary>
        /// Monta um SELECT para o filtro de um campo especifico
        /// </summary>
        /// <param name="fieldName">Nome do Campoa ser filtrado</param>
        /// <param name="fieldValue">Valor a ser filtrado</param>
        /// <returns>String com o SELECT montado</returns>
        public string TSQLSelectByField(string fieldName, string fieldValue)
        {
            return string.Format("{0} WHERE {2}={3}", TSQLSelectAll, fieldName, fieldValue);
        }
        /// <summary>
        /// Monta um SELECT para o filtro de um campo especifico
        /// </summary>
        /// <param name="fieldName">Nome do Campoa ser filtrado</param>
        /// <param name="fieldValue">Valor a ser filtrado</param>
        /// <returns>String com o SELECT montado</returns>
        public string TSQLSelectByField()
        {
            return TSQLSelectAll;
        }

        /// <summary>
        /// Baseado nas propriedades da classe, Monta uma string seprada por virgula com os campos para uso nos SELECT
        /// </summary>
        /// <returns></returns>
        private string GetSelectFields()
        {
            if (_Properties==null)
            {
                GetProperties(thisClass);
            }
            return string.Join("  ,", propString.ToArray());
        }


        public string GetInsertFields()
        {
            GetProperties(thisClass);
            string sql_insert = "";

            foreach (var p in _Properties)
            {
                sql_insert = sql_insert + string.Format(",{0}", getTSQLValue(p: p));
            }
            return sql_insert.Trim().Substring(1);
        }

        public string GetUpdateFields()
        {
            GetProperties(thisClass);
            string sql_update="";

            foreach (var p in _Properties)
            {
                sql_update = sql_update + string.Format(",[{0}]={1}", p.Name, getTSQLValue(p:p));
            }
            return sql_update;
        }

        private object getTSQLValue(Property p)
        {
            {
                var retorn = "";
                switch (p.DataType.ToLower())
                {
                    case "string":
                        retorn = "'" + p.Value?.ToString() + "'";
                        break;
                    default:
                        retorn = p.Value.ToString();
                        break;
                };
                return retorn;
            };
        }


        #endregion

        /// <summary>
        /// Usando Reflections.GetProperties, analisa as informações da Classe obtem os nomes e valores das propriedades
        /// </summary>
        /// <param name="objClass">Instancia da classe</param>
        /// <returns>Lista de Propriedades (Property)</returns>
        internal List<Property> GetProperties(object objClass)
        {
            thisClass = objClass;
            var list = objClass.GetType().GetProperties().ToList();

            Properties = new List<Property>();
            propString = new List<string>();

            var attC = getMapFlag(objClass.GetType());

            list.ForEach(p =>
            {
                if ((p.DeclaringType.Name == thisClass.GetType().Name))
                {
                    //var att = getMapFlag(p.GetType());
                    //Condições para adicionar a propriedade na lista ou nao
                    // 1) A classe tem a marcação e o valor é true e 
                    var att = p.GetType().GetCustomAttributesData();
                    //MappForField vAtt;
                    if (att?.Count >0)
                    {
                        Properties.Add(new Property() { Name = p.Name, Value = getMemberValue(objClass, p), DataType = p.PropertyType.Name });
                        propString.Add(string.Format("[{0}]", p.Name));
                    }

                }
            });
            return Properties;
        }

        private MappForField getMapFlag(dynamic obj)
        {
            var attC = obj.GetCustomAttributesData();
            //MappForField vAtt;
            
            if (attC?.Count >= 1 && attC[0]?.Constructor.DeclaringType.Name == "MappForField")
            {
                var vAtt = obj.GetCustomAttributes(true);
                return vAtt[0];
            }
            else
                //var vAtt = null;
                return null;
        }

        /// <summary>
        /// Obtem o valor de uma propriedade a informar o nome
        /// </summary>
        /// <param name="pName">Nome da Propriedade</param>
        /// <returns>Retorna o conteudo da propriedade</returns>
        public object getValue(string pName)
        {
            var propMember = thisClass.GetType().GetProperty(pName);
            return getMemberValue(thisClass, propMember);
        }
        /// <summary>
        /// Salva um valor na propriedae
        /// </summary>
        /// <param name="pName">Nome da Propriedae</param>
        /// <param name="pValue">Valor a ser salvo</param>
        public void setValue(string pName, object pValue)
        {
            var propMember = thisClass.GetType().GetProperty(pName);
            setMemberValue(thisClass, propMember, pValue);
        }
        /// <summary>
        /// Metodo abstrato que salva o valor proprieamente dito no objeto da classe atraves da função membro.GetValue()
        /// </summary>
        /// <param name="objeto">Instancia do Objeto</param>
        /// <param name="membro">Instancia do membro do objeto</param>
        /// <param name="valor">Valor</param>
        private void setMemberValue(object objeto, System.Reflection.PropertyInfo membro, object valor)
        {
            try
            {
                if (membro.GetGetMethod().IsStatic)
                {
                    membro.SetValue(null, valor, null);
                }
                else
                    membro.SetValue(objeto, valor, null);
            }
            catch (System.Exception ex)
            {
                throw new System.Exception("setMemberValue() - Erro :" + ex.Message);
            }

        }
        /// <summary>
        /// Recupera o valor de um membro de uma classe (Normalmente Propriedade) atraves da função membro.GetValue()
        /// </summary>
        /// <param name="objeto">Instancia do Objeto</param>
        /// <param name="membro">Instancia do membro do objeto</param>
        /// <returns></returns>
        private object getMemberValue(object objeto, System.Reflection.PropertyInfo membro)
        {
            if (membro.GetGetMethod().IsStatic)
            {
                return membro.GetValue(null, null);
            }
            else
                return membro.GetValue(objeto, null);
        }
    }
    /// <summary>
    /// Representa uma Propriedade generica
    /// </summary>
    public class Property
    {
        public string Name { get; set; }
        public object Value { get; set; }
        public string DataType { get; set; }
    }

    /// <summary>
    /// Cria um atributo customizado para marcar uma propriedade como campo mapeado
    /// </summary>
    [System.AttributeUsage(System.AttributeTargets.Property)]
    public sealed class MappForField : System.Attribute
    {
        public bool Map { get; set; }
        public MappForField(bool map)
        {
            this.Map = map;
        }
    }
}
