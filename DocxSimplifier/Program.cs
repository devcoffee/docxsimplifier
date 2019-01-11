/*****************************************************************************
* Produto: DocxSimplifier                                                    *
* Copyright (C) 2018  devCoffee Sistemas de Gestão Integrada                 *
*                                                                            *
* Este arquivo é parte do DocxSimplifier que é software livre; você pode     *
* redistribuí-lo e/ou modificá-lo sob os termos da Licença Pública Geral GNU,*
* conforme publicada pela Free Software Foundation; tanto a versão 3 da      *
* Licença como (a seu critério) qualquer versão mais nova.                   *
*                                                                            *
*                                                                            *
* Este programa é distribuído na expectativa de ser útil, mas SEM            *
* QUALQUER GARANTIA; sem mesmo a garantia implícita de                       *
* COMERCIALIZAÇÃO ou de ADEQUAÇÃO A QUALQUER PROPÓSITO EM                    *
* PARTICULAR. Consulte a Licença Pública Geral GNU para obter mais           *
* detalhes.                                                                  *
*                                                                            *
* Você deve ter recebido uma cópia da Licença Pública Geral GNU              *
* junto com este programa; se não, escreva para a Free Software              *
* Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA                   *
* 02111-1307, USA  ou para devCoffee Sistemas de Gestão Integrada,           *
* Rua Paulo Rebessi 665 - Cidade Jardim - Leme/SP.                           *
 ****************************************************************************/

using System;
using System.IO;


namespace DocxSimplifier
{
    /*
    * @author Pedro Pozzi Ferreira @PozziSan
    * 
    * This is the main class of the application. His function is to get 
    * the commmand line arguments and manage the functions of Util.cs
    */
    class Program
    {
        static readonly string xNamespace = "www.devcoffee.com.br";
        static void Main(string[] args)
        {

            string docName;
            string location;
            string docLocation;
            bool removeStyles = false;

            //Check if there is Command Line Arguments, if not, Print on terminal a manual of the application usage.
            if (args.Length == 0)
            {
                Console.WriteLine("O nome ou localização do Arquivo não foi informado! Uso correto: DocxSimplifier.exe docLocation\\docName.docx. " +
                    "Adicione a flag --removeStyles caso queira remover a formatação do arquivo");
                Environment.Exit(1);
            }
            else if (args.Length > 1)
            {
                //Check if the --removeStyles flag was given
                if (args[1].Equals("--removeStyles"))
                    removeStyles = true;
            }

            //Set Absolute path to the File Location
            docName = args[0];
            location = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
            docLocation = File.Exists(docName) ? docName : string.Format("{0}\\{1}", location, docName);

            Console.WriteLine("Simplificando documento: " + docName);

            try
            {
                //Calls the function SimplifyMarkup. Notice that the Boolean removeStyles is passed as an argument.
                //If there's no error on function, it prints on console a success message
                Util.SimplifyMarkup(docLocation, xNamespace, removeStyles);
                Console.WriteLine("Sucesso ao Simplificar o Arquivo: {0}", docName);
            }
            catch (Exception e)
            {
                //If an error happens, it's displayed on console
                Console.WriteLine(e.ToString());
            }

        }
    }
}
