using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace WordToPdf.Producer.Models
{
    public class WordToPdfModel
    {
        public string Email { get; set; }
        public IFormFile WordFile { get; set; }
    }
}
