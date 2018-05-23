using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace R2Tech_Valida_Arquivos
{
    class Util : Form1
    {
        public void FecharProcesso(string processNome)
        {
            try
            {
                foreach (Process p in Process.GetProcessesByName(processNome))
                {
                    p.Kill();
                }
            }
            catch
            {
                //tempo de esperar do loop
                System.Threading.Thread.Sleep(500);
            }
        }

        //EXTRAIR DATA DO ARQUIVO
        public string ExtrairDataArquivo(string fileName)
        {
            string tst = "";
            string[] dtTrata = fileName.Split(und);
            foreach (string dt in dtTrata)
            {
                if (dt.IndexOf(tt) != -1)
                {
                    tst = dtTrata[0] + "_" + dt;
                    break;
                }
            }
            return tst;
        }

        //EXTRAIR DATA DO ARQUIVO RELATORIO
        public string ExtrairDataArquivoRelatorio(string fileName)
        {
            string tst = "";
            string[] dtTrata = fileName.Split(und);
            foreach (string dt in dtTrata)
            {
                tst = dtTrata[0] + "_" + dtTrata[2];
                break;
            }
            return tst;
        }

        //EXTRAIR NOME_DATA_IDPOS
        public string ExtrairNomeDataIDPOS(string txt)
        {
            string tst = "";
            string[] dtTrata = txt.Split(und);
            foreach (string dt in dtTrata)
            {
                tst = dtTrata[0] + "_" + dtTrata[1] + "_" + dtTrata[2];
                break;
            }
            return tst;
        }

    }
}
