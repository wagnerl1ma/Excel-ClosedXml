using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml;
using System.Net.Security;
using System.Text.Json;

namespace AppExcel
{
    public static class Excel
    {
        #region Inserir Dados no Excel formatado como tabela
        public static void InserirDadosExcelEmail()
        {
            string caminhoDoArquivo = "C:\\Users\\Wagner\\OneDrive\\TesteListaEmail\\ListaEmail.xlsx";
            //var listaAtualExcel = LerExcelEmail("");

            try
            {
                //if (listaAtualExcel != null)
                if (File.Exists(caminhoDoArquivo))
                {
                    using (var workbook = new XLWorkbook(caminhoDoArquivo))
                    {
                        var abaExcel = workbook.Worksheet("ListaEmail"); //Nome da aba do excel
                        //var nonEmptyDataRows = workbook.Worksheet(1).RowsUsed();

                        var dadosEmailExcel = new List<EmailExcel>();
                        dadosEmailExcel.Add(new EmailExcel() { Id = Guid.NewGuid().ToString(), Nome = "Juliana", Email = "teste@teste.com.br", FoiEnviado = "TESTE3" });

                        int numeroDaUltimaLinha = abaExcel.LastRowUsed().RowNumber(); // numero da ultima linha do excel que está preenchida

                        foreach (var item in dadosEmailExcel)
                        {
                            workbook.Table("TB_EMAIL").InsertRowsBelow(1); // insere mais uma linha na tabela depois do ultimo dado
                            numeroDaUltimaLinha++;

                            abaExcel.Cell("A" + numeroDaUltimaLinha).Value = item.Id;
                            abaExcel.Cell("B" + numeroDaUltimaLinha).Value = item.Nome;
                            abaExcel.Cell("C" + numeroDaUltimaLinha).Value = item.Email;
                            abaExcel.Cell("D" + numeroDaUltimaLinha).Value = item.FoiEnviado;
                        }

                        workbook.SaveAs(caminhoDoArquivo);
                    }
                }
                else
                {
                    Console.WriteLine("Lista Vazia");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ERRO AO INSERIR DADOS NO EXCEL! {ex.Message}");
            }
        }
        #endregion

        public static List<EmailExcel> LerExcelEmail(string caminho)
        {
            string caminhoDoArquivo = "C:\\Users\\Wagner\\OneDrive\\TesteListaEmail\\ListaEmail.xlsx";

            if (File.Exists(caminhoDoArquivo))
            {
                List<EmailExcel> lista = new List<EmailExcel>();              //criamos uma lista vazia para receber cada linha
                using (var workbook = new XLWorkbook(caminhoDoArquivo)) // abrimos o objeto do tipo XLWorkbook 
                {
                    var nonEmptyDataRows = workbook.Worksheet(1).RowsUsed();            // obtem apenas as linhas que foram utilizadas da planilha

                    foreach (var dataRow in nonEmptyDataRows)                           // percorremos linha a linha da planilha
                    {
                        if (dataRow.RowNumber() > 1)                                    //obteremos apenas após a linha 1 para não carregar o cabeçalho
                        {
                            var email = new EmailExcel();                           // criamos um objeto para popular com os valores obtidos da linha
                            email.Id = dataRow.Cell(1).Value.ToString();        // obtemos o valor de cada célula pelo seu nº de coluna
                            email.Nome = dataRow.Cell(2).Value.ToString();
                            email.Email = dataRow.Cell(3).Value.ToString();
                            email.FoiEnviado = dataRow.Cell(4).Value.ToString();

                            lista.Add(email);                                         // adicionamos o objeto criado à lista
                        }
                    }
                    Console.WriteLine(JsonSerializer.Serialize(lista));                 // pronto, exibimos a lista em formato Json
                    return lista;
                }
            }
            else
                Console.WriteLine("Arquivo nao encontrado:" + caminhoDoArquivo);
            return null;

        }

        public static void LerExcel(string caminho)
        {
            string caminhoDoArquivo = System.IO.Directory.GetCurrentDirectory() + caminho;

            if (File.Exists(caminhoDoArquivo))
            {
                List<DespesasModel> lista = new List<DespesasModel>();              //criamos uma lista vazia para receber cada linha

                //var workbook = new XLWorkbook(filePathName);  
                using (var workbook = new XLWorkbook(caminhoDoArquivo)) // abrimos o objeto do tipo XLWorkbook 
                {
                    var nonEmptyDataRows = workbook.Worksheet(1).RowsUsed();            // obtem apenas as linhas que foram utilizadas da planilha

                    foreach (var dataRow in nonEmptyDataRows)                           // percorremos linha a linha da planilha
                    {
                        if (dataRow.RowNumber() > 1)                                    //obteremos apenas após a linha 1 para não carregar o cabeçalho
                        {
                            var despesa = new DespesasModel();                          // criamos um objeto para popular com os valores obtidos da linha
                            despesa.Id = Convert.ToInt32(dataRow.Cell(1).Value);        // obtemos o valor de cada célula pelo seu nº de coluna
                            despesa.Fornecedor = dataRow.Cell(2).Value.ToString();
                            despesa.ValorDevido = Convert.ToDecimal(dataRow.Cell(3).Value);
                            despesa.Descricao = dataRow.Cell(7).Value.ToString();

                            DateTime.TryParse(dataRow.Cell(4).Value.ToString(), out DateTime dataVencto);
                            despesa.Vencimento = Convert.ToDateTime(dataVencto.ToString("dd/MM/yyyy")).Date;

                            if (!string.IsNullOrEmpty(dataRow.Cell(5).Value.ToString()))
                            {
                                DateTime.TryParse(dataRow.Cell(5).Value.ToString(), out DateTime dataPagto);
                                despesa.Pagamento = Convert.ToDateTime(dataPagto.ToString("dd/MM/yyyy")).Date;
                            }

                            if (!string.IsNullOrEmpty(dataRow.Cell(6).Value.ToString()))
                                despesa.ValorPago = Convert.ToDecimal(dataRow.Cell(6).Value);

                            lista.Add(despesa);                                         // adicionamos o objeto criado à lista
                        }
                    }
                    Console.WriteLine(JsonSerializer.Serialize(lista));                 // pronto, exibimos a lista em formato Json
                }
            }
            else
                Console.WriteLine("Arquivo nao encontrado:" + caminhoDoArquivo);
        }

        public static void GerarExcel(string nomeAbaExcel, string nomeDoArquivo, ICollection<DespesasModel> lista)
        {

            var nomeArquivoAtualizado = nomeDoArquivo.Replace(" ", "").Trim() + "_" + DateTime.Now.ToString().Replace('/', '-').Replace(" ", "-").Replace(":", "-").Trim() + ".xlsx";

            //string caminhoDoArquivo = System.IO.Directory.GetCurrentDirectory() + "\\" + nomeDoArquivo;
            string caminhoDoArquivo = "C:\\Users\\Wagner\\Desktop\\Gerar e Ler Excel\\AppExcel\\AppExcel\\Planilhas\\" + nomeArquivoAtualizado;

            if (File.Exists(caminhoDoArquivo))
                File.Delete(caminhoDoArquivo);

            using (var workbook = new XLWorkbook())
            {
                var planilha = workbook.Worksheets.Add(nomeAbaExcel);

                int line = 1;
                GerarCabecalho(planilha);
                line++;

                foreach (var item in lista)
                {
                    planilha.Cell("A" + line).Value = item.Id;
                    planilha.Cell("B" + line).Value = item.Fornecedor;
                    planilha.Cell("C" + line).Value = item.ValorDevido;
                    planilha.Cell("D" + line).Value = item.Vencimento;
                    planilha.Cell("E" + line).Value = item.Pagamento;
                    planilha.Cell("F" + line).Value = item.ValorPago;
                    planilha.Cell("G" + line).Value = item.Descricao;
                    line++;
                }
                workbook.SaveAs(caminhoDoArquivo);
            }
        }

        private static void GerarCabecalho(IXLWorksheet worksheet)
        {
            // Neste método geramos o cabeçalho que ficará na primeira linha das colunas
            worksheet.Cell("A1").Value = "Código";
            worksheet.Cell("B1").Value = "Fornecedor";
            worksheet.Cell("C1").Value = "Valor R$";
            worksheet.Cell("D1").Value = "Vencimento";
            worksheet.Cell("E1").Value = "Pagamento";
            worksheet.Cell("F1").Value = "Valor Pago";
            worksheet.Cell("G1").Value = "Descrição";
        }

        public static List<DespesasModel> GerarDados()
        {
            List<DespesasModel> lstDespesa = new List<DespesasModel>();
            lstDespesa.Add(new DespesasModel()
            {
                Id = 1,
                Fornecedor = "FABRICA A",
                ValorDevido = 500,
                Vencimento = DateTime.Today.AddDays(30),
                Descricao = "Conta usada para teste",
                Pagamento = DateTime.Today.AddDays(30),
                ValorPago = 500
            });
            lstDespesa.Add(new DespesasModel()
            {
                Id = 1,
                Fornecedor = "FABRICA B",
                ValorDevido = 600,
                Vencimento = DateTime.Today.AddDays(30),
                Descricao = "Conta usada para teste",
                Pagamento = DateTime.Today.AddDays(30),
                ValorPago = 600
            });
            lstDespesa.Add(new DespesasModel()
            {
                Id = 1,
                Fornecedor = "FABRICA C",
                ValorDevido = 700,
                Vencimento = DateTime.Today.AddDays(30),
                Descricao = "Conta usada para teste",
                Pagamento = DateTime.Today.AddDays(30),
                ValorPago = 700
            });
            lstDespesa.Add(new DespesasModel()
            {
                Id = 1,
                Fornecedor = "FABRICA D",
                ValorDevido = 800,
                Vencimento = DateTime.Today.AddDays(30),
                Descricao = "Conta usada para teste",
                Pagamento = DateTime.Today.AddDays(30),
                ValorPago = 800
            });
            return lstDespesa;
        }
    }
}
