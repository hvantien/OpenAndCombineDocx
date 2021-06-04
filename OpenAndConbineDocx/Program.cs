using OpenXmlPowerTools;
using System;
using System.Collections.Generic;
using System.IO;

namespace OpenAndConbineDocx
{
    class Program
    {
        static void Main(string[] args)
        {
            string outputFile = @"C:\Users\hvant\Documents\doc3.docx";
            List<string> listFileInput = new List<string>() { @"C:\Users\hvant\Documents\doc1.docx", @"C:\Users\hvant\Documents\doc2.docx" };
            List<MemoryStream> documents = new List<MemoryStream>();
            foreach(var item in listFileInput)
            {
                documents.Add(ReadAllBytesToMemoryStream(item));
            }
            var byteOutput = OpenAndCombine(documents);
            File.WriteAllBytes(outputFile, byteOutput); 
        }

        public static MemoryStream ReadAllBytesToMemoryStream(string path)
        {
            byte[] buffer = File.ReadAllBytes(path);
            var destStream = new MemoryStream(buffer.Length);
            destStream.Write(buffer, 0, buffer.Length);
            destStream.Seek(0, SeekOrigin.Begin);
            return destStream;
        }

        public static byte[] OpenAndCombine(List<MemoryStream> documents)
        {
            List<Source> documentBuilderSources = new List<Source>();
            foreach (MemoryStream documentByteArray in documents)
            {
                documentBuilderSources.Add(new Source(new WmlDocument(string.Empty, documentByteArray), true));
            }
            WmlDocument mergedDocument = DocumentBuilder.BuildDocument(documentBuilderSources);
            return mergedDocument.DocumentByteArray;
        }
    }
}
