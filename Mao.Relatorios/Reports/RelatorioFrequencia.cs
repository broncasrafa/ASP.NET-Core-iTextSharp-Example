using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Mao.Relatorios.Classes;
using Mao.Relatorios.Core.PDF;

namespace Mao.Relatorios.Reports
{
    public class RelatorioFrequencia
    {
        private static PdfConverter GetConfigurations()
        {
            PdfConverter pdfConverter = new PdfConverter();
            pdfConverter.DocumentOptions.PageOrientation = PdfPageOrientation.Landscape;
            pdfConverter.DocumentOptions.ShowFooter = true;
            pdfConverter.DocumentOptions.ShowHeader = true;
            pdfConverter.DocumentOptions.ShowPagingOnBottom = true;
            pdfConverter.DocumentOptions.ShowPrintDateTime = true;
            pdfConverter.DocumentOptions.HeaderOptions.DrawHeaderLine = true;
            pdfConverter.DocumentOptions.HeaderOptions.HeaderTitleText = "RELATÓRIO DE FREQUÊNCIA";
            pdfConverter.DocumentOptions.HeaderOptions.HeaderImage = System.Drawing.Image.FromFile($"{Directory.GetCurrentDirectory()}\\images\\fde_logo_2.jpg");
            pdfConverter.DocumentOptions.FooterOptions.DrawFooterLine = true;
            pdfConverter.DocumentOptions.FooterOptions.ShowPageNumber = true;
            pdfConverter.DocumentOptions.DocumentInfo.AuthorName = "Rafael Francisco";
            pdfConverter.DocumentOptions.DocumentInfo.CreatedDate = DateTime.Now;
            pdfConverter.DocumentOptions.DocumentInfo.Title = "Relatório de Frequência";
            pdfConverter.DocumentOptions.DocumentInfo.Subject = "Relatório gerado para testes";
            pdfConverter.DocumentOptions.DocumentInfo.Creator = "Sistema de teste de geração de PDF";

            return pdfConverter;
        }
        private static string GetHtmlContent(List<Frequencia> dados)
        {
            StringBuilder html = new StringBuilder(PdfMethods.LoadHtmlTemplate($"{Directory.GetCurrentDirectory()}\\templates_relatorios\\RelatorioFrequencia_Template.html"));
            StringBuilder resultado = new StringBuilder();

            foreach (var item in dados)
            {
                resultado.AppendLine(@"  <tr>");
                resultado.AppendLine($@"      <td class='valor_item'>{item.Data}</td>");
                resultado.AppendLine($@"      <td class='valor_item'>{item.Entrada}</td>");
                resultado.AppendLine($@"      <td class='valor_item'>{item.Saida}</td>");
                resultado.AppendLine($@"      <td class='valor_item'>{item.TotalHorasDia}</td>");
                resultado.AppendLine($@"      <td class='valor_item' style='text-align: left;'>{item.Tarefa}</td>");
                resultado.AppendLine(@"  </tr>");
            }
            resultado.AppendLine($@"     <tr>
                                            <td colspan='3' class='fundo_cinza bold valor_item'>Total Mês</td>
                                            <td class='fundo_cinza bold valor_item'>{dados.SomarPeriodos(c => c.TotalHorasDia).ToString()}</td>
                                            <td class='fundo_cinza valor_item'></td>
                                         </tr>");

            html = html.Replace("#KEY_RESULT", resultado.ToString());

            return html.ToString();
        }

        public static byte[] Generate(List<Frequencia> dados)
        {
            PdfConverter pdf = GetConfigurations();
            return pdf.RelatorioFrequencia(dados);
        }
        public static byte[] GenerateFromHtml(List<Frequencia> dados)
        {
            string cssStyles = @"                
            .fundo_cinza { background-color: #B5B5B5; }
            .p_init { font-size: 15px; padding-bottom: 30px; } 
            .valor_item { height: 25px; text-align: center; }
            .head { height: 25px; text-align: center; }
            .bold { font-weight: bold; }";

            string html = GetHtmlContent(dados);

            PdfConverter pdf = GetConfigurations();
            return pdf.GeneratePdfBytesFromHtmlString(html, cssStyles);
        }
    }
}
