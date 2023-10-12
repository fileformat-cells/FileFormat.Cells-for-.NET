using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

namespace FileFormat.Cells
{
    

    
    public class Image
    {
        public string Path { get; }
        public Stream Data { get; }
        public string Extension { get; }


        // You can add more properties as needed, such as format, size, etc.

        public Image(string path)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
                throw new ArgumentException("Valid path required", nameof(path));

            Path = path;
            Extension = System.IO.Path.GetExtension(path);
        }

        // Constructor to handle images that originate from streams
        public Image(Stream data, string extension)
        {
            Data = data ?? throw new ArgumentNullException(nameof(data));
            Extension = extension ?? throw new ArgumentNullException(nameof(extension));
        }

    }


    

}

