using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
//REFERENCIAS UTILITARIOS
using Ionic.Zip;

namespace R2Tech_Valida_Arquivos
{
    class TrataArquivos : Form1
    {
        public Boolean VerificarArquivos(string extensao, string keyword)
        {
            try
            {

                //filePath //caminho do arquivo
                //OBTER ARQUIVOS
                if (extensao == "xlsx" && keyword != "")
                {
                    arquivos = Directory.GetFiles(filePath, "*" + keyword + "*." + extensao, SearchOption.TopDirectoryOnly);
                }
                else if(extensao == "xlsx")
                {
                    arquivos = Directory.GetFiles(filePath, "*." + extensao, SearchOption.TopDirectoryOnly);
                }
                else
                {
                    arquivos = Directory.GetFiles(filePath, "*." + extensao, SearchOption.TopDirectoryOnly);
                }
                //VALIDAR RETORNO DE ARQUIVOS LOCALIZADOS
                if (arquivos.Count() > 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                return false;
            }
        }

        //DESCOMPACTAR ARQUIVOS ZIP
        public Boolean DescompactarArquivosZIP(string caminho)
        {
            //verificar arquivos zip na pasta
            if (zips.Count() > 0)
            {
                for (int i = 0; i < zips.Count(); i++)
                {
                    using (ZipFile fileZip = new ZipFile(caminho + zips[i]))
                    {
                        try
                        {
                            fileZip.Password = "Auditeste";
                            fileZip.ExtractAll(caminho);
                            System.Threading.Thread.Sleep(500);
                        }
                        catch(Exception ex)
                        {
                            System.Windows.Forms.MessageBox.Show(ex.Message);
                        }
                    }
                }

                return true;
            }

            return false;
        }

    }
}
