using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Domain
{
    //private string _name;
    //private object _value;
    //public Property() { }
    //public Property(string pName, object pValue) { _name = pName; _value = pValue; }
    //public string Name { get { return _name; } set { _name = value; } }
    //public object Value { get { return _value; } set { _value = value; } }

    public interface IProperty<P>
    {
        string Name { get; set; }
        object value { get; set; }
    }

    /// <summary>
    /// Classe base para uma propriedade generica
    /// </summary>
    public class Property1
    {
        public Property1() { }
        public Property1(string pName, object pValue) { this.Name = pName; this.Value = pValue; }
        public string Name { get; set; }
        public object Value { get; set; }
    }
}