using System;

namespace TinyConfig
{
    class Program
    {
        private static string excelPath = "测试.xlsx";

        static void Main(string[] args)
        {
            var options = new Converter.Options();
            options.excelPath = excelPath;
            options.nameSpace = "";
            options.clientFolder = "./";
            options.serverFolder = "./";
            options.csFolder = "./";

            new Converter(options);
            Console.WriteLine("Convert Success");

            string a = System.IO.File.ReadAllText("./CombatExpDef.csv");
            CombatExpDef.Load(a);
            Console.WriteLine("CombatExpDef : " + CombatExpDef.Values.Count);
            CombatExpDef def = CombatExpDef.Find(2);
            Console.WriteLine("CombatExpDef Level=2: " +  def.Name);

            string b = System.IO.File.ReadAllText("./OtherDef.csv");
            OtherDef.Load(b);
            Console.WriteLine("OtherDef : " + OtherDef.Values.Count);
            OtherDef def2 = OtherDef.Find(1, 2);
            Console.WriteLine("OtherDef id=1, Level=2: " + def2.Name);

            Console.ReadLine();
        }
    }
}
