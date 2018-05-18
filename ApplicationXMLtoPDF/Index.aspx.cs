using RazorEngine;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml.Serialization;
using Pechkin.Synchronized;
using Pechkin;
using System.Drawing.Printing;

namespace ApplicationXMLtoPDF
{
    public partial class Index : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btnProcesar_Click(object sender, EventArgs e)
        {
            Create();
        }

        public void Create()
        {
            Comprobante oComprobante;
            string pathXML = @"E:\ASP NET\ApplicationXMLtoPDF\ApplicationXMLtoPDF\xml\archivoXML4.xml";
            XmlSerializer oSerializer = new XmlSerializer(typeof(Comprobante));
            using (StreamReader reader = new StreamReader(pathXML))
            {
                oComprobante = (Comprobante)oSerializer.Deserialize(reader);

                foreach (var oComplemento in oComprobante.Complemento)
                {
                    foreach (var oComplementoInterior in oComplemento.Any)
                    {
                        if (oComplementoInterior.Name.Contains("TimbreFiscalDigital"))
                        {
                            XmlSerializer oSerializerComplemento = new XmlSerializer(typeof(TimbreFiscalDigital));
                            using (var readerComplemento = new StringReader(oComplementoInterior.OuterXml))
                            {
                                oComprobante.TimbreFiscalDigital = (TimbreFiscalDigital)oSerializerComplemento.Deserialize(readerComplemento);
                            }
                        }
                    }
                }
            }

            //Paso 2 Aplicando Razor y haciendo HTML a PDF

            string path = Server.MapPath("~") + "/";
            string pathHTMLTemp = path + "miHTML.html";//temporal
            string pathHTPlantilla = path + "plantilla.html";
            string sHtml = GetStringOfFile(pathHTPlantilla);
            string resultHtml = "";
            resultHtml = Razor.Parse(sHtml, oComprobante);

            //Creamos el archivo temporal
            File.WriteAllText(pathHTMLTemp, resultHtml);

                       

            GlobalConfig gc = new GlobalConfig();
            gc.SetMargins(new Margins(100, 100, 100, 100))
            .SetDocumentTitle("Test document")
            .SetPaperSize(PaperKind.Letter);

            // Create converter
            IPechkin pechkin = new SynchronizedPechkin(gc);

            // Create document configuration object
            ObjectConfig configuration = new ObjectConfig();


            string HTML_FILEPATH = pathHTMLTemp;

            // and set it up using fluent notation too
            configuration
            .SetAllowLocalContent(true)
            .SetPageUri(@"file:///" + HTML_FILEPATH);

            // Generate the PDF with the given configuration
            // The Convert method will return a Byte Array with the content of the PDF
            // You will need to use another method to save the PDF (mentioned on step #3)
            byte[] pdfContent = pechkin.Convert(configuration);
            ByteArrayToFile(path + "prueba.pdf", pdfContent);
            //eliminamos el archivo temporal
            File.Delete(pathHTMLTemp);
            
        }

        public bool ByteArrayToFile(string _FileName, byte[] _ByteArray)
        {
            try
            {
                // Open file for reading
                FileStream _FileStream = new FileStream(_FileName, FileMode.Create, FileAccess.Write);
                // Writes a block of bytes to this stream using data from  a byte array.
                _FileStream.Write(_ByteArray, 0, _ByteArray.Length);

                // Close file stream
                _FileStream.Close();

                return true;
            }
            catch (Exception _Exception)
            {
                Console.WriteLine("Exception caught in process while trying to save : {0}", _Exception.ToString());
            }

            return false;
        }

        private static string GetStringOfFile(string pathFile)
        {
            string contenido = File.ReadAllText(pathFile);
            return contenido;
        }
    }
}