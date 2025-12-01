using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using ClosedXML.Excel;

namespace ColetaAutomatizadaTG.Services;

public class ExcelExporter : IExcelExporter
{
    public void ExportarCotacoesParaExcel(IEnumerable<CotacaoDto> cotacoes, string caminhoArquivo)
    {
        try
        {
            if (cotacoes is null)
                throw new ArgumentNullException(nameof(cotacoes));

            var diretorio = Path.GetDirectoryName(caminhoArquivo);
            if (!string.IsNullOrWhiteSpace(diretorio) && !Directory.Exists(diretorio))
                Directory.CreateDirectory(diretorio);

            using var workbook = new XLWorkbook();

            var planilha = workbook.Worksheets.Add("Cotações");

            planilha.Cell(1, 1).Value = "Data";
            planilha.Cell(1, 2).Value = "Valor de Compra";
            planilha.Cell(1, 3).Value = "Valor de Venda";
            planilha.Cell(1, 4).Value = "Moeda";

            var cabecalho = planilha.Range(1, 1, 1, 4);
            cabecalho.Style.Font.Bold = true;
            cabecalho.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cabecalho.Style.Fill.BackgroundColor = XLColor.LightGray;

            int linha = 2;

            foreach (var c in cotacoes)
            {
                planilha.Cell(linha, 1).Value = c.Data;

                if (TryParseDecimalPtBr(c.ValorCompra, out var valorCompra))
                {
                    var cellCompra = planilha.Cell(linha, 2);
                    cellCompra.Value = valorCompra;
                    cellCompra.Style.NumberFormat.Format = "R$ #,##0.0000";
                }
                else
                {
                    planilha.Cell(linha, 2).Value = c.ValorCompra;
                }

                if (TryParseDecimalPtBr(c.ValorVenda, out var valorVenda))
                {
                    var cellVenda = planilha.Cell(linha, 3);
                    cellVenda.Value = valorVenda;
                    cellVenda.Style.NumberFormat.Format = "R$ #,##0.0000";
                }
                else
                {
                    planilha.Cell(linha, 3).Value = c.ValorVenda;
                }

                planilha.Cell(linha, 4).Value = c.TipoMoeda;

                linha++;
            }

            planilha.Columns().AdjustToContents();
            planilha.SheetView.FreezeRows(1);

            using var ms = new MemoryStream();
            workbook.SaveAs(ms);
            var novoArquivoBytes = ms.ToArray();
            long novoTamanho = novoArquivoBytes.Length;

            if (File.Exists(caminhoArquivo))
            {
                long tamanhoExistente = new FileInfo(caminhoArquivo).Length;

                if (tamanhoExistente == novoTamanho)
                {
                    throw new IOException($"O arquivo '{caminhoArquivo}' já existe com o mesmo tamanho.");
                }
            }

            File.WriteAllBytes(caminhoArquivo, novoArquivoBytes);
        }
        catch (Exception)
        {
            throw;
        }
    }


    private static bool TryParseDecimalPtBr(string? valor, out decimal resultado)
    {
        resultado = 0m;

        if (string.IsNullOrWhiteSpace(valor))
            return false;

        var texto = valor.Trim();

        if (decimal.TryParse(texto, NumberStyles.Any, new CultureInfo("pt-BR"), out resultado))
            return true;

        if (decimal.TryParse(texto, NumberStyles.Any, CultureInfo.InvariantCulture, out resultado))
            return true;

        return false;
    }
}
