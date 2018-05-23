using Bytescout.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace R2Tech_Valida_Arquivos
{
    class Planilhas : Form1
    {
        public int num = 0;
        public static Spreadsheet arqExcel;
        public static Worksheet planilha;
        public static decimal varUtil;
        public static Util util = new Util();

        //-----------------
        public static List<string> valor = new List<string>();
        public static List<string> pos = new List<string>();
        //public static Dictionary<int, string> dicPosValor = new Dictionary<int, string>();    //DICCIONARIO VALORES POS
        //-----------------

        public Boolean AbrirPlanilha(string filePath, string fileName)
        {
            //verificar arquivo existente
            if (File.Exists(filePath + fileName))
            {
                //abrir planilha
                arqExcel = new Spreadsheet();

                //DECLARANDO CAMINHO E ARQUIVO
                arqExcel.LoadFromFile(filePath + fileName);

                //PLANILHA CONTAGEM E ABERTURA
                num = arqExcel.Workbook.Worksheets.Count;
                if (num == 1)
                {
                    //ABRINDO A PLANILHA
                    var nomePlan = arqExcel.Worksheet(0).Name;
                    planilha = arqExcel.Workbook.Worksheets.ByName(nomePlan);

                    if (ValidarConteudoPlanilha(filePath, fileName) == true)
                    {
                        arqExcel.Dispose();
                        arqExcel.Close();
                        arqExcel = null;
                        nomePlan = null;
                        return true;
                    }
                }   
            }
            else
            {
                if (arqExcel != null)
                {
                    arqExcel.Dispose();
                    arqExcel.Close();
                    arqExcel = null;
                }
                MessageBox.Show("Arquivo Não Localizado" + filePath + fileName);
                return false;
            }
            return false;
        }

        //VALIDAR CONTEUDO PLANILHA MATCH
        public Boolean ValidarConteudoPlanilha(string filePath, string fileName)
        {
            //EXTRAIR NOME DO ARQUIVO
            nomeArquivo = util.ExtrairDataArquivo(fileName);

            //IDENTIFICAR COLUNAS
            row = 0;
            i = 0;
            while (planilha.Cell(row, i).Value != null)
            {
                dicCols.Add(i, planilha.Cell(row, i).Value.ToString());
                i++;
            }

            //SOMAR VALOR COLUNA FIN. VLR. LIQ. E VALOR COLUNA DIFERENÇA
            row = 1;
            i = 0;

            if (idMaquineta == "")
            {
                idMaquineta = planilha.Cell(row, i).Value.ToString(); //INICIALIZA PRIMEIRA LINHA
            }

            //PERCORRER LINHAS
            while (planilha.Cell(row, i).Value != null)
            {
                //PERCORRER LINHAS
                FinVlrLiq = FinVlrLiq + ColetarValor("Fin. Vlr. Liq.", row);
                Diferenca = Diferenca + ColetarValor("Diferença", row);

                //QUANTIFICAR OS IDS MAQUINETA
                if (idMaquineta != planilha.Cell(row, i).Value.ToString())
                {
                    idMaquineta = planilha.Cell(row, i).Value.ToString();
                    qtdeIDS++;
                }
                else
                {
                    if (nomeArquivoAnterior == "")
                    {
                        nomeArquivoAnterior = nomeArquivo;
                    }

                    if (nomeArquivoAnterior != nomeArquivo)
                    {
                        idMaquineta = planilha.Cell(row, i).Value.ToString();
                        qtdeIDS++;
                        nomeArquivoAnterior = nomeArquivo;
                    }
                }

                row++;
            }
            finvlrliq = Convert.ToDouble(FinVlrLiq);
            ValAnterior = FinVlrLiq;

            //VALIDAR SE PASSOU DOS 1% DA SOMA TOTAL
            res = finvlrliq * 0.01;
            Resultado = Convert.ToDecimal(res);

            if (Diferenca > Resultado)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        //REALIZAR SOMA PELOS ID'S MATCHOPB
        public Boolean RealizarSomaIds(string filePath, string fileName)
        {
            //verificar arquivo existente
            if (File.Exists(filePath + fileName))
            {
                //abrir planilha
                arqExcel = new Spreadsheet();

                //DECLARANDO CAMINHO E ARQUIVO
                arqExcel.LoadFromFile(filePath + fileName);

                //PLANILHA CONTAGEM E ABERTURA
                num = arqExcel.Workbook.Worksheets.Count;
                if (num == 1)
                {
                    //ABRINDO A PLANILHA
                    var nomePlan = arqExcel.Worksheet(0).Name;
                    planilha = arqExcel.Workbook.Worksheets.ByName(nomePlan);

                    //EXTRAIR NOME DO ARQUIVO
                    nomeArquivo = util.ExtrairDataArquivo(fileName);
                    //////////////////////OPERACAO DE SOMA DE TODOS OS ID'S

                    //IDENTIFICAR COLUNAS
                    row = 0;
                    i = 0;
                    while (planilha.Cell(row, i).Value != null)
                    {
                        dicCols.Add(i, planilha.Cell(row, i).Value.ToString());
                        i++;
                    }
                    //IDENTIFICAR IDS DA PLANILHA

///////////////////////////////////////////////////
                    //IDENTIFICANDO OS IDS MAQUINETA 
                    row = 1;
                    i = 0;
                    idMaquineta = "";
                    //INICIALIZA PRIMEIRA LINHA
                    if (idMaquineta == "")
                    {
                        idMaquineta = planilha.Cell(row, i).Value.ToString();
                    }

                    //PERCORRER LINHAS
                    while (planilha.Cell(row, i).Value != null)
                    {
                        //QUANTIFICAR OS IDS MAQUINETA
                        if (idMaquineta != planilha.Cell(row, i).Value.ToString())
                        {
                            idMaquineta = planilha.Cell(row, i).Value.ToString();
                            qtdeIDS++;
                        }

                        row++;
                    }
                    //FIM QUANTIFICAR ID'S MAQUINETA
                    ///////////////////////////////////////////////////////////////////////////

                    //COLETAR IDS E JOGAR NO VETOR
                    idMaquineta = "";
                    row = 1;
                    i = 0;
                    idsPOS = new string[qtdeIDS];
                    indexIDS = 0;
                    //for (int j = 0; j < qtdeIDS; j++)
                    //{
                    //    idMaquineta = "";
                    //    row = 1;

                    //INICIALIZA PRIMEIRA LINHA
                    if (idMaquineta == "")
                    {
                        idMaquineta = planilha.Cell(row, i).Value.ToString();
                        idsPOS[indexIDS] = planilha.Cell(row, i).Value.ToString();
                        indexIDS++;
                    }
                    //PERCORRER LINHAS
                    while (planilha.Cell(row, i).Value != null)
                    {
                        //COLETAR IDS MAQUINETA
                        if (idMaquineta != planilha.Cell(row, i).Value.ToString())
                        {
                            idMaquineta = planilha.Cell(row, i).Value.ToString();
                            idsPOS[indexIDS] = planilha.Cell(row, i).Value.ToString();
                            indexIDS++;
                        }
                        row++;
                    }
                    //}

                    //TRATAMENTO IDS MAQUINETA 
                    distincIDSPOS = idsPOS.Distinct().ToArray();


                    ///////////////////////////////////////////////// SOMA DOS VALORES POR IDS MAQUINETA
                    indexIDS = 0;
                    //dicIDSPOS.Clear();
                    //SOMA DOS VALORES 
                    for (int j = 0; j < distincIDSPOS.Count(); j++)
                    {
                        //SOMAR VALOR COLUNA FIN. VLR. LIQ. POR ID'S
                        idMaquineta = "";
                        row = 1;
                        i = 0;
                        somaValPOS = 0;

                        if (idMaquineta == "")
                        {
                            //INICIALIZA PRIMEIRA LINHA
                            idMaquineta = distincIDSPOS[j];
                        }

                        //PERCORRER LINHAS
                        while (planilha.Cell(row, i).Value != null)
                        {
                            //QUANTIFICAR OS IDS MAQUINETA
                            var tst = planilha.Cell(row, i).Value.ToString();
                            if (idMaquineta == planilha.Cell(row, i).Value.ToString())
                            {
                                somaValPOS = somaValPOS + ColetarValor("Fin. Vlr. Liq.", row);
                            }
                            row++;
                        }
                        dicIDSPOS.Add(anteriorIDS, nomeArquivo + "_" + idMaquineta + "_" + somaValPOS);
                        //somaPOS[anteriorIDS] = nomeArquivo + "_" + idMaquineta + "_" + somaValPOS;
                        anteriorIDS++;
                        indexIDS++;
                        
                    }

                    dicCols.Clear();
                    //////////////////////OPERACAO DE SOMA DE TODOS OS ID'S

                    if (arqExcel != null)
                    {
                        arqExcel.Dispose();
                        arqExcel.Close();
                        arqExcel = null;
                        nomePlan = null;
                        return true;
                    }
                }
            }
            else
            {
                if (arqExcel != null)
                {
                    arqExcel.Dispose();
                    arqExcel.Close();
                    arqExcel = null;
                }
                MessageBox.Show("Arquivo Não Localizado" + filePath + fileName);
                return false;
            }
            return false;


        }

        //VALIDAR CONTEUDO PLANILHA RELATORIO
        public Boolean ValidarConteudoPlanilhaRelatorio(string filePath, string fileName)
        {
            //verificar arquivo existente
            if (File.Exists(filePath + fileName))
            {
                nomeArquivo = util.ExtrairDataArquivoRelatorio(fileName);

                //abrir planilha
                arqExcel = new Spreadsheet();

                //DECLARANDO CAMINHO E ARQUIVO
                arqExcel.LoadFromFile(filePath + fileName);

                //PLANILHA CONTAGEM E ABERTURA
                num = arqExcel.Workbook.Worksheets.Count;
                somaValPOS = 0;
                if (num == 3)
                {
                    //ABRINDO A PLANILHA
                    var planGERAL = arqExcel.Worksheet(0).Name;
                    var planPOS = arqExcel.Worksheet(1).Name;
                    var planDETALHE = arqExcel.Worksheet(2).Name;

////////////////////
                    //DADOS DE PLANILHA GERAL
                    planilha = arqExcel.Workbook.Worksheets.ByName(planGERAL);
                    if (planilha != null)
                    {
                        //IDENTIFICAR COLUNAS
                        row = 0;
                        i = 0;
                        while (planilha.Cell(row, i).Value != null)
                        {
                            dicCols.Add(i, planilha.Cell(row, i).Value.ToString());
                            i++;
                        }

                        //COLETAR SALDO INICIAL, FINAL, DISPONIBILIZADO E CANCELAMENTOS
                        row = 1;
                        i = 0;
                        while (planilha.Cell(row, i).Value != null) //PERCORRER LINHAS
                        {
                            //PERCORRER COLUNAS
                            i = 0;
                            while (i < dicCols.Count())
                            {   //SaldoInicial SaldoFinal SaldoAnterior Disponibilizado Cancelamentos
                                if (planilha.Cell(row, i).Value.ToString().ToUpper() == "SALDO INICIAL")
                                {
                                    SaldoInicial = ColetarValor("Saldo", row);
                                }

                                if (planilha.Cell(row, i).Value.ToString().ToUpper() == "DISPONIBILIZADO")
                                {
                                    Disponibilizado = ColetarValor("Valor", row);
                                }

                                if (planilha.Cell(row, i).Value.ToString().ToUpper() == "CANCELAMENTOS")
                                {
                                    Cancelamentos = ColetarValor("Valor", row);
                                }

                                if (planilha.Cell(row, i).Value.ToString().ToUpper() == "SALDO FINAL")
                                {
                                    SaldoFinal = ColetarValor("Saldo", row);
                                }

                                i++;
                            }
                            i = 0;
                            row++;
                        }

                    }
                    else
                    {
                        arqExcel.Dispose();
                        arqExcel.Close();
                        arqExcel = null;
                        MessageBox.Show("Planilha " + planGERAL + "NÃO LOCALIZADA NO ARQUIVO");
                        return false;
                    }

                    dicCols.Clear();
                    ////////////////////////////
                    //DADOS DE PLANILHA POS
                    planilha = arqExcel.Workbook.Worksheets.ByName(planPOS);
                    if (planilha != null)
                    {
                        //IDENTIFICAR COLUNAS
                        row = 0;
                        i = 0;
                        while (planilha.Cell(row, i).Value != null)
                        {
                            dicCols.Add(i, planilha.Cell(row, i).Value.ToString());
                            i++;
                        }

                        //COLETAR SALDO INICIAL, FINAL, DISPONIBILIZADO E CANCELAMENTOS
                        row = 1;
                        i = 0;
                        var ativLoop = 1;
                        var iloop = 0;
                        somaValPOS = 0;
                        //dicPosValor.Clear();

                        //PERCORRER LINHAS
                        while (ativLoop == 1)//planilha.Cell(row, i).Value != null
                        {
                            if (planilha.Cell(row, i).Value == null)
                            {
                                if (iloop >= 2)
                                {
                                    ativLoop = 0;
                                    break;
                                }
                                else
                                {
                                    iloop = 0;
                                    iloop++;
                                    row++;
                                }
                            }
                            //PERCORRER COLUNAS
                            i = 0;

                            //-----------------
                            int colunaPOS = 1; //Coluna do ID POS = segunda coluna posicao 1 //para guardar no dicionario
                            //string[] arrayValorPOS = new string[dicIDSPOS.Count()];
                            //string[] arraySomaValPOS = new string[dicIDSPOS.Count()];
                            //-----------------

                            if (planilha.Cell(row, i).Value != null)
                            {
                                while (i < dicCols.Count())
                                {   
                                    if (planilha.Cell(row, i).Value.ToString().ToUpper() == "DISPONIBILIZADO")
                                    {
                                        somaValPOS = somaValPOS + ColetarValor("Valor", row);
                                    }
                                    if (planilha.Cell(row, i).Value.ToString().ToUpper() == "CANCELAMENTOS")
                                    {
                                        somaValPOS = somaValPOS + ColetarValor("Valor", row);
                                    }
                                    if (planilha.Cell(row, i).Value.ToString().ToUpper() == "CHARGEBACKS")
                                    {
                                        somaValPOS = somaValPOS + ColetarValor("Valor", row);
                                    }
                                    if (planilha.Cell(row, i).Value.ToString().ToUpper() == "OUTROS AJUSTES")
                                    {
                                        somaValPOS = somaValPOS + ColetarValor("Valor", row);
                                    }

                                    //-----------------
                                    //se estiver na linha saldofinal, entao guarda a soma e o id
                                    if (planilha.Cell(row, i).Value.ToString().ToUpper() == "SALDO FINAL")
                                    {
                                        //aqui colocar validacao para caso somaValPOS tiver apenas uma casa depois da virgula, então adicionar um 0, exemplo: 11,5 para 11,50.
                                        int inicioAposVirgulaSomaValPOS = somaValPOS.ToString().IndexOf(",");
                                        string somaValPOSAposVirgula = somaValPOS.ToString().Substring(inicioAposVirgulaSomaValPOS + 1);

                                        string strSomaPosValor = "";

                                        //Primeiro tratamento
                                        if (somaValPOSAposVirgula.Length == 1)
                                        {
                                            strSomaPosValor = somaValPOS.ToString() + "0";
                                            dicPosValor.Add(indexIDS, nomeArquivo + "_" + planilha.Cell(row, colunaPOS).Value.ToString() + "_" + strSomaPosValor);
                                        }
                                        else
                                        {
                                            dicPosValor.Add(indexIDS, nomeArquivo + "_" + planilha.Cell(row, colunaPOS).Value.ToString() + "_" + somaValPOS.ToString());
                                        }

                                        indexIDS++;
                                        //pos.Add(planilha.Cell(row, colunaPOS).Value.ToString());
                                        //valor.Add(somaValPOS.ToString());
                                        somaValPOS = 0;
                                    }
                                    //-----------------

                                    i++;
                                }

                            }
                            else
                            {
                                iloop++;
                            }

                            ////-----------------

                            //for (int b = 0; b < valor.Count(); b++)
                            //{
                            //    //dicPosValor.Add(b+1, arrayValorPOS[b] + "_" + arraySomaValPOS[b]);
                            //    dicPosValor.Add(b+1, pos[b] + "_" + valor[b]);
                            //}
                            ////-----------------

                            i = 0;
                            row++;
                        }

                        dicCols.Clear();
                        dicValues.Clear();
                        arqExcel.Dispose();
                        arqExcel.Close();
                        arqExcel = null;
                        return true;
                    }
                    else
                    {
                        arqExcel.Dispose();
                        arqExcel.Close();
                        arqExcel = null;
                        MessageBox.Show("Planilha " + planPOS + "NÃO LOCALIZADA NO ARQUIVO");
                        return false;
                    }


                    ////////////////////////////



                }
            }
            else
            {
                arqExcel.Dispose();
                arqExcel.Close();
                arqExcel = null;

                MessageBox.Show("Arquivo Não Localizado" + filePath + fileName);
                return false;
            }

            return false;
        }

        //COLETAR VALOR PELA COLUNA
        public decimal ColetarValor(string nameCol, int row)
        {
            foreach (KeyValuePair<int, string> item in dicCols)
            {
                if (item.Value.ToUpper() == nameCol.ToUpper())
                {
                    varUtil = (decimal)planilha.Cell(row, item.Key).ValueAsDouble;
                    break;
                }
            }
            return varUtil;
        }

    }
}
