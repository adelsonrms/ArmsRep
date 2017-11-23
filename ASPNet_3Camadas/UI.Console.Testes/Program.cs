using System.Linq;

namespace UI.Console.Testes
{
    class Program
    {
        static void Main(string[] args)
        {

            var mc = new DTO.Cliente();

            mc.setValue("Nome", "Adelson");
            System.Console.WriteLine("------------------------------\n");
            System.Console.WriteLine("Instrução Select para a classe\n");
            System.Console.Write(mc.TSQLSelectByID(1));
            System.Console.ReadKey();

            var ListProp = mc.GetType().GetProperties();

            ListProp.ToList().ForEach(p => System.Console.WriteLine(buscaInfo(mc, p)));


            


            System.Console.WriteLine("Altera os valores");
            System.Console.ReadKey();

            for (int i = 0; i < ListProp.ToList().Count; i++)
            {
                alterarValor(objeto:mc,  membro: ListProp[i], valor:i);
            }

            System.Console.WriteLine("Atualiza com os novos valores");
            System.Console.ReadKey();

            //ListProp = mc.GetType().GetProperties();
            ListProp.ToList().ForEach(p => System.Console.WriteLine(buscaInfo(mc, p)));

            
            System.Console.WriteLine("\nPressione uma tecla pra continuar");
            System.Console.ReadKey();

        }

        private static void alterarValor(object objeto,  System.Reflection.PropertyInfo membro, object valor)
        {
            if (membro.GetMethod.IsStatic)
            {
                membro.SetValue(null, valor);
            }
            else
                membro.SetValue(objeto, valor);
        }


        private static object pegaValor(object objeto, System.Reflection.PropertyInfo membro)
        {
            if (membro.GetMethod.IsStatic)
            {
               return membro.GetValue(null);
            }
            else
                return membro.GetValue(objeto);
        }  



        private static string buscaInfo(object objeto , System.Reflection.PropertyInfo membro)
        {
            return string.Format("Nome : {0} / Valor Atual : {3} / Tipo : {1}  / Estatico {2}", membro.Name, membro.PropertyType.Name, membro.GetMethod.IsStatic, pegaValor(objeto,membro));
        }

        

    }

    class MyClass
    {
        public static int Nome { get; set; }
        public int ID { get; set; }
    }
}
