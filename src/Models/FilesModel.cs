using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace OfficeAndPdfConverter.Models
{
    public class FilesModel
    {
        public int FileId { get; set; } = 0;

        public string Name { get; set; } = "";

        public string Path { get; set; } = "";

        public List<FilesModel> Files { get; set; } = new List<FilesModel>();

        public string Inofmation { get; set; }

        public int Type { get; set; } //1 pdf 2 word 3 excel 4 powerpoint

    }
}
