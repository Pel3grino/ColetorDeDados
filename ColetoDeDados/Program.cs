using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using System.Globalization;

namespace ColetorDeDados
{
    class Program
    {
        static async Task Main(string[] args)
        {
            Console.WriteLine(@"Digite o caminho onde está salvo o xlsx: ");
            var caminhoPasta = Console.ReadLine();
            Console.WriteLine(@"Digite somente o nome do arquivo (Sem .xlsx): ");
            var arquivoNome = Console.ReadLine();


            var xlsx = new XLWorkbook($"{caminhoPasta}\\{arquivoNome}.xlsx");
            var planilha = xlsx.Worksheets.FirstOrDefault();
            var totalLinhas = planilha.Rows().Count();
            
            var fullPath = $@"{caminhoPasta}\LerNoCPlus5_{DateTime.Now.ToString("ddMMyyyy_hhmmss")}.txt";
            
            List<Produto> list = new List<Produto>();

            for (int l = 2; l <= totalLinhas; l++) // primeira linha é o cabecalho
            {
                var produto = new Produto();
                produto.Codigo = (planilha.Cell($"A{l}").Value.ToString());
                produto.Quantidade = planilha.Cell($"B{l}").GetDouble();
                if (produto.Quantidade <0 )
                {
                    produto.Quantidade = 0;
                }
                produto.TextoColetor = produto.Codigo.PadRight(20).Substring(0,20) + produto.Quantidade.ToString(CultureInfo.InvariantCulture);
                list.Add(produto);
               
            }

            using (var result = new StreamWriter(fullPath.ToString()))
                
            {
                foreach (var lista in list)
                {

                    result.WriteLine(lista.TextoColetor);

                }
                Console.WriteLine($"Arquivo gerado em {fullPath}");
                Console.WriteLine("Pressione qualquer tecla para sair!");
                
                Console.ReadKey();

            }
        }
    }
}






