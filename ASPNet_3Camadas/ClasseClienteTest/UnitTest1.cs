using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using DTO;

namespace ClasseClienteTest
{
    [TestClass]
    public class EntityController_Test
    {
        [TestMethod]
        public void EntityController_TSQLSelectAll()
        {
            var cli = new Cliente();
            int count = cli.Properties.Count;
            var countVirgula = cli.TSQLSelectByField().Split(char.Parse(","));
            var msg = (count == countVirgula.Length ? "OK, a quantidade de propriedades é coerentes" : "Numeros de propriedades não compativel com os campos mapeados");
            Assert.AreEqual(count, countVirgula.Length, msg);
        }

        [TestMethod]
        public void EntityController_TSQLInsert()
        {
            var cli = new Cliente();
            int count = cli.Properties.Count;
            var countVirgula = cli.GetInsertFields();
            var msg = (count == countVirgula.Length ? "OK, a quantidade de propriedades é coerentes" : "Numeros de propriedades não compativel com os campos mapeados");
            Assert.AreEqual(count, countVirgula.Length, msg);
        }

    }
}
