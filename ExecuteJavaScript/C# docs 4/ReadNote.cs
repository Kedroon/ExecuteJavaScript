using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;

namespace PMMBotMySql
{
    public class ReadNote
    {
        string source;
        Uri urlofnote;
        HtmlDocument doc;

        public ReadNote(string src, Uri url)
        {
            source = src;
            urlofnote = url;
        }

        public void StartAnalysis()
        {
            MySqlConnection connection = new MySqlConnection();
            connection.ConnectionString = "Server=localhost;Database=notas;Uid=root;Pwd=;";
           // OleDbConnection connection = new OleDbConnection();
           // connection.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\migue\OneDrive\Documentos\Notas.accdb;
//Persist Security Info=False;";
            string query;

            doc = new HtmlDocument();
            doc.LoadHtml(source.Replace("&nbsp;", ""));
            string CNPJPrestador = "";
            string nfe = "";
            string rps = "";
            string dis = "";
            string valorliquido = "";
            string valorservico = "";
            string ISSQNRetido = "";
            string CODServico = "";
            string NFeSub = "";
            string DataHoraEmissao = "";
            string Competencia = "";
            string CODVerificacao = "";
            string CNPJTomador = "";
            string RazaoSocialNome = "";
            string CIA = "";
            string DescontoIncondicionado = "";
            string DescontoCondicionado = "";
            string RetençõesFederais = "";
            string OutrasRetenções = "";
            string PIS = "";
            string COFINS = "";
            string IR = "";
            string INSS = "";
            string CSLL = "";


            try //Try CNPJ/CPF Prestador
            {
                CNPJPrestador = doc.DocumentNode.SelectSingleNode("//span[starts-with(@style,'position:absolute;left:129px;top:176px;')]").SelectSingleNode(".//*").InnerHtml;
                Console.WriteLine("CNPJ/CPF Prestador: " + CNPJPrestador);
            }
            catch (Exception err)
            {
                Console.WriteLine(err.Message);
                Console.WriteLine("cade o CNPJ Prestador");

            }

            try //Try nota fiscal eletronica
            {
                nfe = doc.DocumentNode.SelectSingleNode("//span[starts-with(@style,'position:absolute;left:502px;top:52px;')]").SelectSingleNode(".//*").InnerHtml;

                Console.WriteLine("NFe: " + nfe);
            }
            catch (Exception)
            {
                nfe = "Não tem nota fiscal???????";
                Console.WriteLine("Não tem nota fiscal???????");

            }


            try //Try Discriminacao
            {
                dis = doc.DocumentNode.SelectSingleNode("//span[starts-with(@style,'position:absolute;left:12px;top:335px;')]").SelectSingleNode(".//*").InnerHtml.Replace("<br>", " ").Replace("'", "''");
                Console.WriteLine("Discriminacao: " + dis);
            }
            catch (Exception)
            {
                dis = "Não possui Discriminação do Serviço";
                Console.WriteLine("Não possui Discriminação");
            }


            try //Try RPS
            {
                rps = doc.DocumentNode.SelectSingleNode("//span[starts-with(@style,'position:absolute;left:124px;top:102px;')]").SelectSingleNode(".//*").InnerHtml;
                Console.WriteLine("RPS: " + rps);

            }
            catch (Exception)
            {
                rps = "Não possui RPS";
                Console.WriteLine("Não possui RPS");
            }

            //Try Valor Liquido

            valorliquido = returnValueofElementDetalhamento("(=) Valor Líquido R$", "Não possui valor liquido");
            //valorliquido = doc.DocumentNode.SelectSingleNode("//span[starts-with(@style,'position:absolute;left:135px;top:690px;')]").SelectSingleNode(".//*").InnerHtml;
            Console.WriteLine("Valor liquido: " + valorliquido);


            //Try Valor Servico

            valorservico = returnValueofElementDetalhamento("Valor do Serviço  R$", "Não possui valor do serviço");
            //valorservico = doc.DocumentNode.SelectSingleNode("//span[starts-with(@style,'position:absolute;left:135px;top:570px;')]").SelectSingleNode(".//*").InnerHtml;
            Console.WriteLine("Valor Servico: " + valorservico);




            //Try ISSQN Retido
            ISSQNRetido = returnValueofElementDetalhamento("(-) ISSQN Retido", "Não possui ISSQN Retido");
            //ISSQNRetido = doc.DocumentNode.SelectSingleNode("//span[starts-with(@style,'position:absolute;left:135px;top:670px;')]").SelectSingleNode(".//*").InnerHtml;
            Console.WriteLine("ISSQN Retido: " + ISSQNRetido);




            //Try Codigo do Servico
            CODServico = returnValueofElemenetCodigoServiço("Código do Serviço / Atividade", "Não possui codigo de serviço");
            //CODServico = doc.DocumentNode.SelectSingleNode("//span[starts-with(@style,'position:absolute;left:12px;top:450px;')]").SelectSingleNode(".//*").InnerHtml;
            //int indexEnd = CODServico.IndexOf("-");
            //CODServico = CODServico.Substring(0, indexEnd - 1);
            Console.WriteLine("Codigo do Servico: " + CODServico);




            try //Try NFe Substituido
            {
                NFeSub = doc.DocumentNode.SelectSingleNode("//span[starts-with(@style,'position:absolute;left:319px;top:102px;')]").SelectSingleNode(".//*").InnerHtml;
                Console.WriteLine("NFe Substituido: " + NFeSub);

            }
            catch (Exception)
            {
                NFeSub = "Nao possui Nfe Substituido";
                Console.WriteLine("Não possui NFe Substituido");
            }

            try //Try Data e Hora
            {
                DataHoraEmissao = doc.DocumentNode.SelectSingleNode("//span[starts-with(@style,'position:absolute;left:124px;top:82px;')]").SelectSingleNode(".//*").InnerHtml;
                Console.WriteLine("Data e Hora de Emissao: " + DataHoraEmissao);

            }
            catch (Exception)
            {
                DataHoraEmissao = "Nao possui Data e Hora de Emissao";
                Console.WriteLine("Não possui Data e Hora de Emissao");
            }

            try //Try Competencia
            {
                Competencia = doc.DocumentNode.SelectSingleNode("//span[starts-with(@style,'position:absolute;left:319px;top:82px;')]").SelectSingleNode(".//*").InnerHtml;
                Console.WriteLine("Competencia: " + Competencia);

            }
            catch (Exception)
            {
                Competencia = "Nao possui Competencia";
                Console.WriteLine("Não possui Competencia");
            }

            try //Try Codigo de Verificação
            {
                CODVerificacao = doc.DocumentNode.SelectSingleNode("//span[starts-with(@style,'position:absolute;left:470px;top:82px;')]").SelectSingleNode(".//*").InnerHtml;
                Console.WriteLine("Codigo de Verificacao: " + CODVerificacao);

            }
            catch (Exception)
            {
                CODVerificacao = "Nao possui Codigo de Verificacao";
                Console.WriteLine("Não possui Codigo de Verificacao");
            }

            try //Try CNPJ do Tomador
            {
                CNPJTomador = doc.DocumentNode.SelectSingleNode("//span[starts-with(@style,'position:absolute;left:59px;top:264px;')]").SelectSingleNode(".//*").InnerHtml;
                Console.WriteLine("CNPJ/CPF Tomador: " + CNPJTomador);

            }
            catch (Exception)
            {
                CNPJTomador = "Nao possui CNPJ Tomador";
                Console.WriteLine("Não possui CNPJ Tomador");
            }

            try //Try CNPJ do RazaoSocialNome
            {
                RazaoSocialNome = doc.DocumentNode.SelectSingleNode("//span[starts-with(@style,'position:absolute;left:175px;top:142px;')]").SelectSingleNode(".//*").InnerHtml;
                Console.WriteLine("Razao Social Nome: " + RazaoSocialNome);

            }
            catch (Exception)
            {
                RazaoSocialNome = "Nao possui Razao Social Nome";
                Console.WriteLine("Não possui Razao Social Nome");
            }

            DescontoIncondicionado = returnValueofElementDetalhamento("(-) Desconto Incondicionado", "Não possui Desconto Incondicionado");
            Console.WriteLine(DescontoIncondicionado);

            DescontoCondicionado = returnValueofElementDetalhamento("(-) Desconto Condicionado", "Não possui Desconto Condicionado");
            Console.WriteLine(DescontoCondicionado);

            RetençõesFederais = returnValueofElementDetalhamento("(-) Retenções Federais", "Não possui Retenções Federais");
            Console.WriteLine(RetençõesFederais);

            OutrasRetenções = returnValueofElementDetalhamento("(-) Outras Retenções", "Não possui Outras Retenções");
            Console.WriteLine(OutrasRetenções);

            PIS = returnValueofElementTributos("PIS(R$)", "73", "Não possui PIS");
            Console.WriteLine(PIS);

            COFINS = returnValueofElementTributos("COFINS(R$)", "185", "Não possui COFINS");
            Console.WriteLine(COFINS);

            IR = returnValueofElementTributos("IR(R$)", "297", "Não possui IR");
            Console.WriteLine(IR);

            INSS = returnValueofElementTributos("INSS(R$)", "409", "Não possui INSS");
            Console.WriteLine(INSS);

            CSLL = returnValueofElementTributos("CSLL(R$)", "528", "Não possui CSLL");
            Console.WriteLine(CSLL);


            if (CNPJPrestador == "04.335.535/0002-55")  //Insert BD SuperTerminais Table
            {
                SuperTerminais superterminais = new SuperTerminais(dis, nfe);
                if (superterminais.BeginAnalysis())
                {
                //insert no banco
                InsertSuperTerminais:
                    try
                    {
                        connection.Open();
                        MySqlCommand command = new MySqlCommand();
                        command.Connection = connection;
                        query = "insert into Notas (NFe, RPS , DiscriminacaodoServico , ValorLiquido , ValorServico , ISSQNRetino , CODServico , NFeSub , DataHoraEmissao , Competencia , CODVerificacao , CNPJPrestador , CNPJTomador , RazaoSocialNomePrestador , DescontoIncondicionado , DescontoCondicionado , RetencoesFederais , OutrasRetencoes , PIS , COFINS , IR , INSS , CSLL , URL) values ('" + nfe + "','" + rps + "','" + dis + "','" + valorliquido + "','" + valorservico + "','" + ISSQNRetido + "','" + CODServico + "','" + NFeSub + "','" + DataHoraEmissao + "','" + Competencia + "','" + CODVerificacao + "','" + CNPJPrestador + "','" + CNPJTomador + "','" + RazaoSocialNome + "','" + DescontoIncondicionado + "','" + DescontoCondicionado + "','" + RetençõesFederais + "','" + OutrasRetenções + "','" + PIS + "','" + COFINS + "','" + IR + "','" + INSS + "','" + CSLL + "','" + urlofnote + "')";
                        command.CommandText = query;
                        command.ExecuteNonQuery();
                        connection.Close();
                    }
                    catch (Exception err)
                    {
                        connection.Close();
                        Console.WriteLine(nfe + " - " + err.Message);
                        goto InsertSuperTerminais;
                    }

                }
                else
                {
                    Console.WriteLine("DI ZOADA");
                }
            }

            else if (CNPJPrestador == "04.694.548/0001-30")  //Insert BD Aurora Table
            {

                AuroraEadi auroraeadi = new AuroraEadi(dis, nfe);
                if (auroraeadi.BeginAnalysis())
                {
                InsertAurora:
                    try
                    {
                        connection.Open();
                        MySqlCommand command = new MySqlCommand();
                        command.Connection = connection;
                        query = "insert into Notas (NFe, RPS , DiscriminacaodoServico , ValorLiquido , ValorServico , ISSQNRetino , CODServico , NFeSub , DataHoraEmissao , Competencia , CODVerificacao , CNPJPrestador , CNPJTomador , RazaoSocialNomePrestador , DescontoIncondicionado , DescontoCondicionado , RetencoesFederais , OutrasRetencoes , PIS , COFINS , IR , INSS , CSLL , URL) values ('" + nfe + "','" + rps + "','" + dis + "','" + valorliquido + "','" + valorservico + "','" + ISSQNRetido + "','" + CODServico + "','" + NFeSub + "','" + DataHoraEmissao + "','" + Competencia + "','" + CODVerificacao + "','" + CNPJPrestador + "','" + CNPJTomador + "','" + RazaoSocialNome + "','" + DescontoIncondicionado + "','" + DescontoCondicionado + "','" + RetençõesFederais + "','" + OutrasRetenções + "','" + PIS + "','" + COFINS + "','" + IR + "','" + INSS + "','" + CSLL + "','" + urlofnote + "')";
                        command.CommandText = query;
                        command.ExecuteNonQuery();
                        connection.Close();
                    }
                    catch (Exception err)
                    {
                        connection.Close();
                        Console.WriteLine(nfe + " - " + err.Message);
                        goto InsertAurora;
                    }

                }
                else
                {
                    Console.WriteLine("Colocaram um navio no meio da avenida!");
                }


            }

            else if (CNPJPrestador == "84.098.383/0001-72")  //Insert BD Chibatao Table
            {
                Chibatao chibatao = new Chibatao(dis, nfe);
                if (chibatao.BeginAnalysis())
                {
                InsertChibatao:
                    try
                    {
                        connection.Open();
                        MySqlCommand command = new MySqlCommand();
                        command.Connection = connection;
                        query = "insert into Notas (NFe, RPS , DiscriminacaodoServico , ValorLiquido , ValorServico , ISSQNRetino , CODServico , NFeSub , DataHoraEmissao , Competencia , CODVerificacao , CNPJPrestador , CNPJTomador , RazaoSocialNomePrestador , DescontoIncondicionado , DescontoCondicionado , RetencoesFederais , OutrasRetencoes , PIS , COFINS , IR , INSS , CSLL , URL) values ('" + nfe + "','" + rps + "','" + dis + "','" + valorliquido + "','" + valorservico + "','" + ISSQNRetido + "','" + CODServico + "','" + NFeSub + "','" + DataHoraEmissao + "','" + Competencia + "','" + CODVerificacao + "','" + CNPJPrestador + "','" + CNPJTomador + "','" + RazaoSocialNome + "','" + DescontoIncondicionado + "','" + DescontoCondicionado + "','" + RetençõesFederais + "','" + OutrasRetenções + "','" + PIS + "','" + COFINS + "','" + IR + "','" + INSS + "','" + CSLL + "','" + urlofnote + "')";
                        command.CommandText = query;
                        command.ExecuteNonQuery();
                        connection.Close();
                    }
                    catch (Exception err)
                    {
                        connection.Close();
                        Console.WriteLine(nfe + " - " + err.Message);
                        goto InsertChibatao;
                    }

                }
                else
                {
                    Console.WriteLine("Colocaram um navio no meio da avenida!");
                }
            }

            else
            {
            InsertWhatever:
                try
                {
                    connection.Open();
                    MySqlCommand command = new MySqlCommand();
                    command.Connection = connection;
                    query = "insert into Notas (NFe, RPS , DiscriminacaodoServico , ValorLiquido , ValorServico , ISSQNRetino , CODServico , NFeSub , DataHoraEmissao , Competencia , CODVerificacao , CNPJPrestador , CNPJTomador , RazaoSocialNomePrestador , DescontoIncondicionado , DescontoCondicionado , RetencoesFederais , OutrasRetencoes , PIS , COFINS , IR , INSS , CSLL , URL) values ('" + nfe + "','" + rps + "','" + dis + "','" + valorliquido + "','" + valorservico + "','" + ISSQNRetido + "','" + CODServico + "','" + NFeSub + "','" + DataHoraEmissao + "','" + Competencia + "','" + CODVerificacao + "','" + CNPJPrestador + "','" + CNPJTomador + "','" + RazaoSocialNome + "','" + DescontoIncondicionado + "','" + DescontoCondicionado + "','" + RetençõesFederais + "','" + OutrasRetenções + "','" + PIS + "','" + COFINS + "','" + IR + "','" + INSS + "','" + CSLL + "','" + urlofnote + "')";
                    command.CommandText = query;
                    command.ExecuteNonQuery();
                    connection.Close();
                }
                catch (Exception err)
                {
                    connection.Close();
                    Console.WriteLine(nfe + " - " + dis + " - " + err.Message);
                    goto InsertWhatever;
                }
            }



        }

        string returnValueofElementDetalhamento(string searchValue, string failure)
        {
            string temp;
            try
            {
                temp = doc.DocumentNode.SelectSingleNode("//span[.='" + searchValue + "']").OuterHtml;
                int indextop = temp.IndexOf("top") + 4;
                temp = temp.Substring(indextop, temp.Substring(indextop).IndexOf("px;"));
                temp = doc.DocumentNode.SelectSingleNode(@"//span[starts-with(@style,'position:absolute;left:135px;top:" + temp + "px;')]").InnerText;
            }
            catch (Exception)
            {
                temp = failure;
            }

            return temp;
        }

        string returnValueofElementTributos(string searchValue, string left, string failure)
        {
            string temp;
            try
            {
                temp = doc.DocumentNode.SelectSingleNode("//span[.='" + searchValue + "']").OuterHtml;
                int indextop = temp.IndexOf("top") + 4;
                temp = temp.Substring(indextop, temp.Substring(indextop).IndexOf("px;"));
                temp = doc.DocumentNode.SelectSingleNode(@"//span[starts-with(@style,'position:absolute;left:" + left + "px;top:" + temp + "px;')]").InnerText;
            }
            catch (Exception)
            {

                temp = failure;
            }

            return temp;
        }

        string returnValueofElemenetCodigoServiço(string searchValue, string failure)
        {
            string temp;
            try
            {
                temp = doc.DocumentNode.SelectSingleNode("//span[.='" + searchValue + "']").OuterHtml;
                int indextop = temp.IndexOf("top") + 4;
                temp = temp.Substring(indextop, temp.Substring(indextop).IndexOf("px;"));
                int i = int.Parse(temp) + 20;
                temp = doc.DocumentNode.SelectSingleNode(@"//span[starts-with(@style,'position:absolute;left:12px;top:" + i + "px;')]").InnerText;
                int indexEnd = temp.IndexOf("-");
                temp = temp.Substring(0, indexEnd - 1);
            }
            catch (Exception)
            {

                temp = failure;
            }
            return temp;
        }

    }

}