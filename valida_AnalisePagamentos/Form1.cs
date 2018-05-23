using Bytescout.Spreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.IO;

namespace R2Tech_Valida_Arquivos
{
    public partial class Form1 : Form
    {
        //VARIAVEIS PLANILHAS
        public static int iTest = 0;
        public static int i = 0;
        public static int row = 1;
        public static decimal FinVlrLiq = 0, Diferenca = 0, Resultado = 0, ValAnterior = 0; //variaveis planilha match
        public static decimal SaldoInicial = 0, SaldoFinal, SaldoAnterior = 0, Disponibilizado = 0, Cancelamentos = 0; //variaveis planilha relatorio movimentacao
        public static Dictionary<int, string> dicCols = new Dictionary<int, string>();
        public static Dictionary<string, string> dicValues = new Dictionary<string, string>();
        public static Dictionary<int, string> dicIDSPOS = new Dictionary<int, string>();    //DICCIONARIO SOMA VALORES MATCHOPB
        //public static Dictionary<int, string> dicVALREL = new Dictionary<int, string>();    //DICCIONARIO VALORES POS RELATORIO
        public static Dictionary<int, string> dicPosValor = new Dictionary<int, string>();  //DICCIONARIO VALORES POS RELATORIO

        public static double finvlrliq = 0, res = 0;
        public static decimal somaValPOS = 0;
        private static Planilhas plan = new Planilhas();
        private static TrataArquivos trataArquivos = new TrataArquivos();
        private static Util util = new Util();

        //public static Excel.Workbook arqExcel;
        //public static Excel.Worksheet planilha;

        //VARIAVEIS TRATAMENTO DE ARQUIVOS
        public static string filePath;
        public static string[] arquivos;        // RETORNO DA LISTA DE ARQUIVOS LOCOLIZADOS NA BUSCA
        public static string[] zips;            // RETORNO ARQUIVOS ZIP LOCALIZADOS
        public static string[] excel;           // PLANILHAS MATCHOPB
        public static string[] pdf;             // ARQUIVOS PDF
        public static string[] planMovi;        // PLANILHAS DE RELATORIO DE MOVIMENTACAO
        public static string[] movPOS;          // SOMA DOS VALORES FIN. MAQUINETA POR IDENTIFICADOR
        public static string[] somaPOS;         // SOMA DOS VALORES POR ID PEGANDO A DATA DE ARQUIVO
        public static string[] idsPOS;
        public static string[] dataLista;       // DATA LISTA DE ARQUIVOS
        public static string[] distincIDSPOS;
        public static char und = '_';
        public static char tt = '-';
        public static int qtdeIDS = 1;
        public static int indexIDS = 0;
        public static int anteriorIDS = 0;
        public static string idMaquineta = "";
        public static string nomeArquivo = "";
        public static string nomeArquivoAnterior = "";

        //VARIAVEIS LOG
        public static string arquivoLog = "\\log_" + DateTime.Now.ToString("yyyy-MM-dd") + ".txt";
        public static string nomeArquivoLog = arquivoLog;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //Ao Abrir com parâmetro pelo CMD

            txtCaminhoDiretorio.Text = Program.valorDoParamentro;
            //string caminho = txtCaminhoDiretorio.Text;

            if (Program.valorDoParamentro != "")
            {

                //System.Threading.Thread.Sleep(3000);

                //Match();
                //Relatorio();
                //ValidaPDFCasasDecimais();
                //GerarLog();
                //Application.Exit();
            }
        }

        // BUSCAR ARQUIVOS ZIPS
        private void btnProcurar_Click(object sender, EventArgs e)
        {
            util.FecharProcesso("EXCEL");
            //ZERAR VARIAVEIS 
            SaldoAnterior = 0;
            SaldoInicial = 0;
            SaldoFinal = 0;
            Disponibilizado = 0;
            Cancelamentos = 0;
            //ZERAR VARIAVEIS

            //LIMPAR NOMES DOS ARQUIVOS
            excel = null;
            planMovi = null;
            pdf = null;
            dataLista = null;
            qtdeIDS = 1;
            idMaquineta = "";
            nomeArquivo = "";
            nomeArquivoAnterior = "";

            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            folderDlg.ShowNewFolderButton = true;

            // Show the FolderBrowserDialog.

            DialogResult result = folderDlg.ShowDialog();

            if (result == DialogResult.OK)
            {
                txtCaminhoDiretorio.Text = folderDlg.SelectedPath;
                Environment.SpecialFolder root = folderDlg.RootFolder;

                if (Program.valorDoParamentro != "")
                {
                    txtCaminhoDiretorio.Text = Program.valorDoParamentro;
                }


                ListaLog.Items.Add(txtCaminhoDiretorio.Text);
                ListaLog.Items.Add("");
                filePath = txtCaminhoDiretorio.Text + @"\";

                //LISTAR ARQUIVOS ZIP
                if (trataArquivos.VerificarArquivos("zip", "") == true)
                {
                    if (arquivos.Count() > 0)
                    {
                        zips = new string[arquivos.Count()];
                        dataLista = new string[arquivos.Count()];

                        ListaLog.Items.Add("ARQUIVOS ZIP ENCONTRADOS");
                        for (int i = 0; i < arquivos.Count(); i++)
                        {
                            ListaLog.Items.Add(arquivos[i]);
                            zips[i] = arquivos[i].Replace(filePath, "");
                            //EXTRAIR DATA DOS ARQUIVOS
                            ListarDataArquivos(zips[i], i);
                        }
                    }
                }
                else
                {
                    ListaLog.Items.Add("ARQUIVOS ZIP NÃO LOCALIZADOS");
                    ListaLog.Items.Add("");
                }

                //trataArquivos = null;
            }
        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }

        private void ListaLog_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        //PROCESSAR PLANILHAS RELATORIOS
        private void Relatorio()
        {
            ListaLog.Items.Add("");
            string caminho = txtCaminhoDiretorio.Text;
            //ARQUIVOS XLSX NAO ZIPADOS
            if (trataArquivos.VerificarArquivos("xlsx", "Relatorio") == true)
            {
                if (arquivos.Count() > 0)
                {
                    planMovi = new string[arquivos.Count()];
                    //ListaLog.Items.Add("");//"ARQUIVOS EXCEL ENCONTRADOS");
                    for (int i = 0; i < arquivos.Count(); i++)
                    {
                        //ListaLog.Items.Add(arquivos[i]);
                        planMovi[i] = arquivos[i].Replace(filePath, "");
                    }
                    //VALIDAR APOS DESCOMPACTACAO A QUANTIDADE DE ARQUIVOS
                    if (planMovi.Count() == arquivos.Count())
                    {
                        ListaLog.Items.Add(planMovi.Count() + " ARQUIVOS EXCEL RELATORIOS LOCALIZADOS");
                    }
                    else
                    {
                        ListaLog.Items.Add(planMovi.Count() + " ARQUIVOS LOCALIZADOS QUANTIDADE EXCEL'S NÃO CONFERE FAVOR REVISAR!");
                    }
                }
            }
            else
            {
                ListaLog.Items.Add("");
                ListaLog.Items.Add("NENHUM ARQUIVO EXCEL RELATORIO FOI LOCALIZADOS");
                ListaLog.Items.Add("");
            }

            //VALIDAR CONTEUDO PLANILHA RELATORIO
            if (planMovi != null)
            {
                indexIDS = 0;
                for (int x = 0; x < planMovi.Count(); x++)
                {
                    //VALIDAR CONTEUDO PLANILHA
                    if (plan.ValidarConteudoPlanilhaRelatorio(filePath, planMovi[x]) == true)
                    {
                        ListaLog.Items.Add("");
                        ListaLog.Items.Add(planMovi[x]);
                        if (SaldoAnterior == 0)
                        {
                            ListaLog.Items.Add("SALDO INICIAL: " + SaldoInicial);
                        }
                        else
                        {
                            ListaLog.Items.Add("SALDO INICIAL: " + SaldoInicial + " = " + "SALDO FINAL ANTERIOR: " + SaldoAnterior);
                        }                     
                        ListaLog.Items.Add("DISPONIBILIZADOS: " + Disponibilizado);
                        ListaLog.Items.Add("CANCELAMENTOS: " + Cancelamentos);
                        ListaLog.Items.Add("SALDO FINAL: " + SaldoFinal);

                        if (SaldoAnterior == SaldoInicial)
                        {
                            ListaLog.Items.Add("RESULTADO DA PLANILHA: OK");
                        }
                        else
                        {
                            if (SaldoAnterior != 0)
                            {
                                ListaLog.Items.Add("NOK");
                                ListaLog.Items.Add("SALDO FINAL ANTERIOR COM ERRO: " + SaldoAnterior);
                            }
                        }

                        SaldoAnterior = SaldoFinal;
                        //ListaLog.Items.Add("RESULTADO PLANILHA OK");
                    }
                    else
                    {
                        ListaLog.Items.Add("");
                        ListaLog.Items.Add(planMovi[x]);
                        ListaLog.Items.Add("RESULTADO PLANILHA COM ERRO!");
                    }

                    SaldoInicial = 0;
                    SaldoFinal = 0;
                    Disponibilizado = 0;
                    Cancelamentos = 0;

                }

                var valValidar = "";
                var valPOSValor = "";
                ListaLog.Items.Add("");
                ListaLog.Items.Add("VALIDAR VALOR PLANILHAS MATCHOPB COM VALORES RELATORIO");
                //CRIAR VALIDAÇÃO DOS DICIONARIOS
                for (int z = 0; z < dicIDSPOS.Count(); z++)
                {
                    valValidar = dicIDSPOS[z];
                    valValidar = valValidar.Replace("-", "");
                    valValidar = util.ExtrairNomeDataIDPOS(valValidar);
                    for (int y = 0; y < dicPosValor.Count(); y++)
                    {
                        valPOSValor = dicPosValor[y];
                        valPOSValor = util.ExtrairNomeDataIDPOS(valPOSValor);
                        if (valValidar == valPOSValor)
                        {
                            valValidar = dicIDSPOS[z];
                            valValidar = valValidar.Replace("-", "");
                            if (valValidar == dicPosValor[y])
                            {
                                ListaLog.Items.Add("");
                                ListaLog.Items.Add("LOCALIZADO VALORES IGUAIS >>>");
                                ListaLog.Items.Add("MATHOPB: " + valValidar + " - " + "RELATORIO: " + dicPosValor[y]);
                                break;
                            }
                            else
                            {
                                ListaLog.Items.Add("");
                                ListaLog.Items.Add("NOK");
                                ListaLog.Items.Add("LOCALIZADO VALORES DIFERENTES >>>");
                                ListaLog.Items.Add("MATHOPB: " + valValidar + " - " + "RELATORIO: " + dicPosValor[y]);
                                break;
                            }
                        }
                    }
                    valValidar = "";
                    valPOSValor = "";
                }
                valValidar = null;
                valPOSValor = null;
                dicPosValor.Clear();
                dicIDSPOS.Clear();




            }

        }
        
        //DESCOMPACTAR E VALIDAR QUANTIDADE DE ARQUIVOS DESCOMPACTADOS
        private void Descompactar_Click(object sender, EventArgs e)
        {
            try
            {

                loader.Visible = true;
                loader.Minimum = 0;
                loader.Maximum = zips.Count() * 2;
                loader.Step = 1;

                for (int x = 0; x < zips.Count() * 2; x++)
                {
                    loader.Value = x;
                    loader.PerformStep();

                    if (x == 0)
                    {   //CHAMAR METODO PARA DECOMPACTAR ZIPS
                        Thread secProccess = new Thread(novaTarefa);
                        secProccess.Start();
                        secProccess.Join();
                    }
                    System.Threading.Thread.Sleep(200);
                }

                //ARQUIVOS XLSX ZIPADOS
                if (trataArquivos.VerificarArquivos("xlsx", "Match") == true)
                {
                    if (arquivos.Count() > 0)
                    {
                        excel = new string[arquivos.Count()];
                        ListaLog.Items.Add("");//"ARQUIVOS EXCEL ENCONTRADOS");
                        for (int i = 0; i < arquivos.Count(); i++)
                        {
                            ListaLog.Items.Add(arquivos[i]);
                            excel[i] = arquivos[i].Replace(filePath, "");
                        }
                        //VALIDAR APOS DESCOMPACTACAO A QUANTIDADE DE ARQUIVOS
                        if (excel.Count() == arquivos.Count())
                        {
                            ListaLog.Items.Add(excel.Count() + " ARQUIVOS EXCEL MATCHOPB LOCALIZADOS");
                        }
                        else
                        {
                            ListaLog.Items.Add(excel.Count() + "ARQUIVOS LOCALIZADOS QUANTIDADE EXCEL'S NÃO CONFERE FAVOR REVISAR!");
                        }
                    }
                }
                else
                {
                    ListaLog.Items.Add("");
                    ListaLog.Items.Add("NENHUM ARQUIVO EXCEL MATCHOPB FOI LOCALIZADOS");
                    ListaLog.Items.Add("");
                }
                //ARQUIVOS XLSX NAO ZIPADOS
                if (trataArquivos.VerificarArquivos("xlsx", "Relatorio") == true)
                {
                    if (arquivos.Count() > 0)
                    {
                        planMovi = new string[arquivos.Count()];
                        ListaLog.Items.Add("");//"ARQUIVOS EXCEL ENCONTRADOS");
                        for (int i = 0; i < arquivos.Count(); i++)
                        {
                            ListaLog.Items.Add(arquivos[i]);
                            planMovi[i] = arquivos[i].Replace(filePath, "");
                        }
                        //VALIDAR APOS DESCOMPACTACAO A QUANTIDADE DE ARQUIVOS
                        if (planMovi.Count() == arquivos.Count())
                        {
                            ListaLog.Items.Add(planMovi.Count() + " ARQUIVOS EXCEL RELATORIO LOCALIZADOS");
                        }
                        else
                        {
                            ListaLog.Items.Add(planMovi.Count() + "ARQUIVOS LOCALIZADOS QUANTIDADE EXCEL'S NÃO CONFERE FAVOR REVISAR!");
                        }
                    }
                }
                else
                {
                    ListaLog.Items.Add("");
                    ListaLog.Items.Add("NENHUM ARQUIVO EXCEL RELATORIO FOI LOCALIZADOS");
                    ListaLog.Items.Add("");
                }

                //ARQUIVOS PDF ZIPADOS
                if (trataArquivos.VerificarArquivos("pdf", "") == true)
                {
                    if (arquivos.Count() > 0)
                    {
                        pdf = new string[arquivos.Count()];
                        ListaLog.Items.Add("");//"ARQUIVOS PDF ENCONTRADOS");
                        for (int i = 0; i < arquivos.Count(); i++)
                        {
                            ListaLog.Items.Add(arquivos[i]);
                            pdf[i] = arquivos[i].Replace(filePath, "");
                        }
                        //VALIDAR APOS DESCOMPACTACAO A QUANTIDADE DE ARQUIVOS
                        if (pdf.Count() == arquivos.Count())
                        {
                            ListaLog.Items.Add(pdf.Count() + "ARQUIVOS PDF LOCALIZADOS");
                        }
                        else
                        {
                            ListaLog.Items.Add(pdf.Count() + "ARQUIVOS LOCALIZADOS QUANTIDADE DE PDF'S NÃO CONFERE FAVOR REVISAR!");
                        }
                    }
                }
                else
                {
                    ListaLog.Items.Add("");
                    ListaLog.Items.Add("ARQUIVOS PDF NÃO LOCALIZADOS");
                    ListaLog.Items.Add("");
                }

                if (loader.Value == loader.Maximum)
                {
                    loader.Visible = false;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnProcessar_Click(object sender, EventArgs e)
        {
            ListaLog.ResetText();
            //Adicionar métodos aqui
            ValidaPDFCasasDecimais();
            Match();
            Relatorio();
            GerarLog();
        }

        //PROCESSAR PLANILHAS
        private void Match()
        {
            anteriorIDS = 0;
            ListaLog.Items.Add("");

            //string caminho = txtCaminhoDiretorio.Text;

            //ARQUIVOS XLSX ZIPADOS
            if (trataArquivos.VerificarArquivos("xlsx", "Match") == true)
            {
                if (arquivos.Count() > 0)
                {
                    excel = new string[arquivos.Count()];
                    //ListaLog.Items.Add("");//"ARQUIVOS EXCEL ENCONTRADOS");
                    for (int i = 0; i < arquivos.Count(); i++)
                    {
                    //    ListaLog.Items.Add(arquivos[i]);
                        excel[i] = arquivos[i].Replace(filePath, "");
                    }

                    //VALIDAR APOS DESCOMPACTACAO A QUANTIDADE DE ARQUIVOS
                    if (excel.Count() == arquivos.Count())
                    {
                        ListaLog.Items.Add(excel.Count() + " ARQUIVOS EXCEL MATCHOPB LOCALIZADOS");
                    }
                    else
                    {
                        ListaLog.Items.Add(excel.Count() + " ARQUIVOS LOCALIZADOS QUANTIDADE EXCEL'S NÃO CONFERE FAVOR REVISAR!");
                    }
                }
            }
            else
            {
                ListaLog.Items.Add("");
                ListaLog.Items.Add("NENHUM ARQUIVO EXCEL MATCHOPB FOI LOCALIZADO");
                ListaLog.Items.Add("");
            }

            //ARQUIVOS XLSX NAO ZIPADOS
            if (trataArquivos.VerificarArquivos("xlsx", "Relatorio") == true)
            {
                if (arquivos.Count() > 0)
                {
                    planMovi = new string[arquivos.Count()];
                    //ListaLog.Items.Add("");//"ARQUIVOS EXCEL ENCONTRADOS");
                    for (int i = 0; i < arquivos.Count(); i++)
                    {
                        //ListaLog.Items.Add(arquivos[i]);
                        planMovi[i] = arquivos[i].Replace(filePath, "");
                    }
                    //VALIDAR APOS DESCOMPACTACAO A QUANTIDADE DE ARQUIVOS
                    if (planMovi.Count() == arquivos.Count())
                    {
                        ListaLog.Items.Add(planMovi.Count() + " ARQUIVOS EXCEL RELATORIOS LOCALIZADOS");
                    }
                    else
                    {
                        ListaLog.Items.Add(planMovi.Count() + " ARQUIVOS LOCALIZADOS QUANTIDADE EXCEL'S NÃO CONFERE FAVOR REVISAR!");
                    }
                }
            }
            else
            {
                ListaLog.Items.Add("");
                ListaLog.Items.Add("NENHUM ARQUIVO EXCEL RELATORIO FOI LOCALIZADOS");
                ListaLog.Items.Add("");
            }

            //ARQUIVOS PDF ZIPADOS
            if (trataArquivos.VerificarArquivos("pdf", "") == true)
            {
                if (arquivos.Count() > 0)
                {
                    pdf = new string[arquivos.Count()];
                    //ListaLog.Items.Add("");//"ARQUIVOS PDF ENCONTRADOS");
                    for (int i = 0; i < arquivos.Count(); i++)
                    {
                        //ListaLog.Items.Add(arquivos[i]);
                        pdf[i] = arquivos[i].Replace(filePath, "");
                    }
                    //VALIDAR APOS DESCOMPACTACAO A QUANTIDADE DE ARQUIVOS
                    if (pdf.Count() == arquivos.Count())
                    {
                        ListaLog.Items.Add(pdf.Count() + " ARQUIVOS PDF LOCALIZADOS");
                    }
                    else
                    {
                        ListaLog.Items.Add(pdf.Count() + "ARQUIVOS LOCALIZADOS QUANTIDADE DE PDF'S NÃO CONFERE FAVOR REVISAR!");
                    }
                }
            }
            else
            {
                ListaLog.Items.Add("");
                ListaLog.Items.Add("ARQUIVOS PDF NÃO LOCALIZADOS");
                ListaLog.Items.Add("");
            }

            //VALIDAR PLANILHAS MATCHOPB E SEUS CONTEUDOS
            if (excel != null)
            {
                for (int x = 0; x < excel.Count(); x++)
                {
                    //VALIDAR CONTEUDO PLANILHA
                    if (plan.AbrirPlanilha(filePath, excel[x]) == true)
                    {
                        ListaLog.Items.Add("");
                        ListaLog.Items.Add(excel[x]);
                        ListaLog.Items.Add("Valor soma Coluna Fin. Vlr. Liq. " + String.Format("{0:0.##}", FinVlrLiq));
                        ListaLog.Items.Add("Valor soma Coluna Diferença " + String.Format("{0:0.##}", Diferenca));
                        ListaLog.Items.Add("1% da soma Coluna Fin. Vlr. Liq.  " + String.Format("{0:0.##}", Resultado));
                        ListaLog.Items.Add("Operação: Se a soma de Diferença ultrapassa 1% de Fin. Vlr. Liq.");
                        ListaLog.Items.Add("Soma Diferença " + String.Format("{0:0.##}", Diferenca) + " Não é maior que 1% " + String.Format("{0:0.##}", Resultado));
                        ListaLog.Items.Add("RESULTADO PLANILHA OK");
                    }
                    else
                    {
                        ListaLog.Items.Add("");
                        ListaLog.Items.Add(excel[x]);
                        ListaLog.Items.Add("Valor soma Coluna Fin. Vlr. Liq. " + String.Format("{0:0.##}", FinVlrLiq));
                        ListaLog.Items.Add("Valor soma Coluna Diferença " + String.Format("{0:0.##}", Diferenca));
                        ListaLog.Items.Add("1% da soma Coluna Fin. Vlr. Liq.  " + String.Format("{0:0.##}", Resultado));
                        ListaLog.Items.Add("Operação: Se a soma de Diferença ultrapassa 1% de Fin. Vlr. Liq.");
                        ListaLog.Items.Add("Soma Diferença " + String.Format("{0:0.##}", Diferenca) + " É maior que 1% " + String.Format("{0:0.##}", Resultado));
                        ListaLog.Items.Add("RESULTADO PLANILHA COM ERRO!");
                    }

                    dicCols.Clear();
                    idMaquineta = "";
                    FinVlrLiq = 0;
                    Diferenca = 0;
                    Resultado = 0;
                }

                //DECLARANDO VETOR COM A QUANTIDADE DE IDS LOCALIZADOS DAS PLANILHAS
                somaPOS = new string[qtdeIDS];
                //EFETUAR SOMA DOS ID PARA CADA PLANILHA
                for (int x = 0; x < excel.Count(); x++)
                {
                    qtdeIDS = 1;
                    //if (plan.RealizarSomaIds(filePath, excel[x]) == true)
                    //{
                    //}
                    plan.RealizarSomaIds(filePath, excel[x]);
                }

            }
           
        }

        //VALIDAR PDF's
        private void ValidaPDFCasasDecimais()
        {

            string caminho = txtCaminhoDiretorio.Text;

            int cont = 1;

            //VALIDAR CASAS DECIMAIS NOS ARQUIVOS PDF
            if (trataArquivos.VerificarArquivos("pdf", "") == true)
            {
                bool saiu = false;

                if (arquivos.Count() > 0)
                {
                    pdf = new string[arquivos.Count()];

                    for (int i = 0; i < arquivos.Count(); i++)
                    {
                        //ListaLog.Items.Add(arquivos[i]);
                        pdf[i] = arquivos[i].Replace(filePath, "");
                    }

                    ListaLog.Items.Add("");//"ARQUIVOS PDF ENCONTRADOS");
                    for (int i = 0; i < pdf.Count(); i++)
                    {

                        saiu = false;

                        //using (PdfReader leitor = new PdfReader(filePath + pdf[i]))
                        using (PdfReader leitor = new PdfReader(arquivos[i]))
                        {
                            var texto = new StringBuilder();

                            //valida apenas primeira pagina
                            for (int j = 1; j <= 1; j++)
                            {
                                texto.Append(PdfTextExtractor.GetTextFromPage(leitor, j));
                                string[] words = Convert.ToString(texto).Split(' ');

                                if (saiu == true)
                                {
                                    break;
                                }

                                foreach (string word in words)
                                {
                                    if (word.Contains(','))
                                    {

                                        int inicioAposVirgula = word.IndexOf(",");
                                        string wordAposVirgula = word.Substring(inicioAposVirgula + 1);


                                        //Primeiro tratamento
                                        if (wordAposVirgula.Length == 2 || word == "0,000000")
                                        {
                                            //ListaLog.Items.Add("OK");
                                        }
                                        else
                                        {
                                            //Trata adicionando apenas numeros e virgula na string
                                            char[] letra = word.ToCharArray();
                                            string novaWord = "";

                                            for (int c = 0; c < letra.Length; c++)
                                            {
                                                if (Char.IsNumber(letra[c]) || letra[c] == ',')
                                                {
                                                    novaWord += letra[c];
                                                }

                                            }

                                            //Segundo tratamento
                                            int inicioNovaWordAposVirgula = novaWord.IndexOf(",");
                                            string novaWordAposVirgula = novaWord.Substring(inicioNovaWordAposVirgula + 1);

                                            //se quantidade de numeros após virgula for igual a 2 ok
                                            if (novaWordAposVirgula.Length != 2) //|| novaWordAposVirgula.Length == 2)
                                            {                                                
                                                //Console.WriteLine("nok");
                                                ListaLog.Items.Add("");
                                                ListaLog.Items.Add("ERRO NO PDF - MAIS DE DUAS CASAS DECIMAIS EM VALORES MONETÁRIOS");
                                                ListaLog.Items.Add(pdf[i]);
                                                ListaLog.Items.Add("NOK");
                                                ListaLog.Items.Add("Contém Virgula: " + novaWord);
                                                ListaLog.Items.Add("quantidade de casas após virgula: " + novaWordAposVirgula.Length);

                                                //limpando variaveis
                                                novaWord = "";
                                                inicioNovaWordAposVirgula = 0;
                                                novaWordAposVirgula = "";

                                                cont++;

                                                saiu = true;
                                                break;
                                            }
                                            else
                                            {
                                                ListaLog.Items.Add("");
                                                ListaLog.Items.Add(pdf[i]);
                                                ListaLog.Items.Add("RESULTADO ARQUIVO PDF OK");
                                                ListaLog.Items.Add("Contém Virgula: " + novaWord);
                                                ListaLog.Items.Add("quantidade de casas após virgula: " + novaWordAposVirgula.Length);

                                                //limpando variaveis
                                                novaWord = "";
                                                inicioNovaWordAposVirgula = 0;
                                                novaWordAposVirgula = "";

                                                cont++;

                                                saiu = true;
                                                break;
                                            }
                                        }
                                    }

                                }
                            }
                        }

                        //ListaLog.Items.Add(arquivos[i]);
                        pdf[i] = arquivos[i].Replace(filePath, "");
                    }

                }
            }
            else
            {
                ListaLog.Items.Add("ARQUIVOS PDF NÃO LOCALIZADOS");
                ListaLog.Items.Add("");
            }

        }

        static void novaTarefa()
        {
            trataArquivos.DescompactarArquivosZIP(filePath);
        }

        private static void ListarDataArquivos(string data, int i)
        {
            string[] dtTrata = data.Split(und);
            foreach (string dt in dtTrata)
            {
                
                if (dt.IndexOf(tt) != -1)
                {
                    dataLista[i] = dtTrata[0] + "_" + dt;
                    break;
                }
            }

        }

        //--------------------------------------------------------------------------------
        // Inicio Criar Log
        //--------------------------------------------------------------------------------

        //Método principal da função criar log
        public void GerarLog()
        {

            string[] arry = new string[ListaLog.Items.Count];

            arry = get();

            //Cria arquivo log
            string caminho = txtCaminhoDiretorio.Text;
            IniciaLog();

            //Adicionando dados da ListBox no arquivo log.txt
            using (StreamWriter w = File.AppendText(caminho + nomeArquivoLog))
            {
                foreach (var item in arry)
                {
                    Log(item, w);
                }
            }

        }

        public void IniciaLog()
        {
            string caminho = txtCaminhoDiretorio.Text;
            try
            {
                //Abrir o arquivo
                StreamWriter log = new StreamWriter(caminho + nomeArquivoLog, true, Encoding.ASCII);

                log.WriteLine("\r\nLog Entry: ");
                log.WriteLine("{0} {1}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString());

                //Fecha o arquivo
                log.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception: " + e.Message);
            }
        }

        public static void Log(string logMessage, TextWriter w)
        {
            w.WriteLine(logMessage);
        }

        public string[] get()
        {
            string[] arr = new string[ListaLog.Items.Count];
            for (int i = 0; i < ListaLog.Items.Count; i++)
            {
                arr[i] = ListaLog.Items[i].ToString();
            }
            return arr;
        }

        //--------------------------------------------------------------------------------
        // Fim Criar Log
        //--------------------------------------------------------------------------------

    }
}
