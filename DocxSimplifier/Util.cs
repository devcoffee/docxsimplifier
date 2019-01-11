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

using OpenXmlPowerTools;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using NPOI.XWPF.UserModel;

namespace DocxSimplifier
{
    /*
    * @author Pedro Pozzi Ferreira @pozzisan
    */
    class Util
    {
        private static object TransformToSimpleXml(XNode node, string defaultParagraphStyleId, string z)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                if (element.Name == W.document)
                    return new XElement(z + "document",
                        new XAttribute(XNamespace.Xmlns + "w", z),
                        element.Element(W.body).Elements()
                            .Select(e => TransformToSimpleXml(e, defaultParagraphStyleId, z)));
                if (element.Name == W.p)
                {
                    string styleId = (string)element.Elements(W.pPr)
                        .Elements(W.pStyle).Attributes(W.val).FirstOrDefault();
                    if (styleId == null)
                        styleId = defaultParagraphStyleId;
                    return new XElement(z + "p",
                        new XAttribute("style", styleId),
                        element.LogicalChildrenContent(W.r).Elements(W.t).Select(t => (string)t)
                            .StringConcatenate());
                }
                if (element.Name == W.sdt)
                    return new XElement(z + "contentControl",
                        new XAttribute("tag", (string)element.Elements(W.sdtPr)
                            .Elements(W.tag).Attributes(W.val).FirstOrDefault()),
                        element.Elements(W.sdtContent).Elements()
                            .Select(e => TransformToSimpleXml(e, defaultParagraphStyleId, z)));
                return null;
            }
            return node;
        }

        private static void ReWriteDocument(string docLocation, XElement simplerXml)
        {
            File.Delete(docLocation);

            using (FileStream fileStream = new FileStream(docLocation, FileMode.Create, FileAccess.Write))
            {
                XWPFDocument newWordDoc = new XWPFDocument();

                foreach (XElement paragraph in simplerXml.Elements())
                {
                    XWPFParagraph newDocParagraph = newWordDoc.CreateParagraph();
                    newDocParagraph.Alignment = ParagraphAlignment.LEFT;
                    XWPFRun newDocRun = newDocParagraph.CreateRun();
                    newDocRun.FontFamily = "Arial";
                    newDocRun.FontSize = 12;
                    newDocRun.IsBold = false;
                    newDocRun.SetText(paragraph.Value);
                }

                newWordDoc.Write(fileStream);
                newWordDoc.Close();
            }

        }

        public static void SimplifyMarkup(string docLocation, string z, bool formatDocument)
        {
            try
            {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(docLocation, true))
                {
                    RevisionAccepter.AcceptRevisions(wordDoc);

                    SimplifyMarkupSettings settings = new SimplifyMarkupSettings
                    {
                        RemoveComments = true,
                        RemoveContentControls = true,
                        RemoveEndAndFootNotes = true,
                        RemoveFieldCodes = false,
                        RemoveLastRenderedPageBreak = true,
                        RemovePermissions = true,
                        RemoveProof = true,
                        RemoveRsidInfo = true,
                        RemoveSmartTags = true,
                        RemoveSoftHyphens = true,
                        ReplaceTabsWithSpaces = true,
                        NormalizeXml = false,
                        RemoveWebHidden = true,
                        RemoveMarkupForDocumentComparison = true,

                    };

                    MarkupSimplifier.SimplifyMarkup(wordDoc, settings);

                    string defaultParagraphStyleId = wordDoc.MainDocumentPart
                       .StyleDefinitionsPart.GetXDocument().Root.Elements(W.style)
                       .Where(e => (string)e.Attribute(W.type) == "paragraph" &&
                           (string)e.Attribute(W._default) == "1")
                       .Select(s => (string)s.Attribute(W.styleId))
                       .FirstOrDefault();
                    XElement simplerXml = (XElement)TransformToSimpleXml(
                        wordDoc.MainDocumentPart.GetXDocument().Root,
                        defaultParagraphStyleId, z);
                    Console.WriteLine(simplerXml);

                    wordDoc.Save();
                    wordDoc.Close();
                    if (formatDocument)
                    {
                        Console.WriteLine("Eu resetaria o Documento agora");
                        try
                        {
                            ReWriteDocument(docLocation, simplerXml);
                            Console.WriteLine("Sucesso ao Reformatar o documento!");

                        }
                        catch (Exception e)
                        {
                            throw new Exception(string.Format("Erro ao Reformatar o Arquivo: {0}", e.ToString()));
                        }
                    }

                }


            }
            catch (Exception e)
            {
                throw new Exception(string.Format("Não foi Possível simplificar o Arquivo. Erro: {0}", e.ToString()));
            }

        }
    }
}
