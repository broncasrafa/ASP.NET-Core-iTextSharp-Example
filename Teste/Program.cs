using System;
using System.Collections.Generic;
using System.IO;
using Mao.Relatorios.Classes;
using Mao.Relatorios.Reports;

namespace Teste
{
    class Program
    {
        private static string _PATH = @"C:\temp";


        static void Main(string[] args)
        {
            var listaFrequencia = Frequencia.GetListaFrequencia();
            TesteNormal(listaFrequencia);
            TesteHtml(listaFrequencia);

            var prestacaoServico = PrestacaoServicos.GetPrestacaoServicos();
            TesteNormal2(prestacaoServico);
            TesteHtml2(prestacaoServico);

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        private static void TesteNormal(List<Frequencia> listaFrequencia)
        {
            Console.WriteLine($"Gerando arquivo NORMAL.... {DateTime.Now}");
            
            var fileBytes = RelatorioFrequencia.Generate(listaFrequencia);

            var filename = $"teste_normal_{Guid.NewGuid()}.pdf";

            File.WriteAllBytes($@"{_PATH}\{filename}", fileBytes);

            Console.WriteLine($"PDF salvo com sucesso - {filename} .... {DateTime.Now}");
        }

        private static void TesteHtml(List<Frequencia> listaFrequencia)
        {
            Console.WriteLine($"Gerando arquivo HTML.... {DateTime.Now}");

            var fileBytes = RelatorioFrequencia.GenerateFromHtml(listaFrequencia);

            var filename = $"teste_html_{Guid.NewGuid()}.pdf";

            File.WriteAllBytes($@"{_PATH}\{filename}", fileBytes);

            Console.WriteLine($"PDF salvo com sucesso - {filename} .... {DateTime.Now}");
        }


        private static void TesteNormal2(PrestacaoServicos dados)
        {
            Console.WriteLine($"Gerando arquivo NORMAL.... {DateTime.Now}");

            var fileBytes = RelatorioPrestacaoServico.Generate(dados);

            var filename = $"teste_normal_{Guid.NewGuid()}.pdf";

            File.WriteAllBytes($@"{_PATH}\{filename}", fileBytes);

            Console.WriteLine($"PDF salvo com sucesso - {filename} .... {DateTime.Now}");
        }
        private static void TesteHtml2(PrestacaoServicos dados)
        {
            Console.WriteLine($"Gerando arquivo HTML.... {DateTime.Now}");

            var fileBytes = RelatorioPrestacaoServico.GenerateFromHtml(dados);

            var filename = $"teste_html_{Guid.NewGuid()}.pdf";

            File.WriteAllBytes($@"{_PATH}\{filename}", fileBytes);

            Console.WriteLine($"PDF salvo com sucesso - {filename} .... {DateTime.Now}");
        }
    }
}
