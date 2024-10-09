using APILib;

namespace API.Models
{
    public class FileResult
    {
        public Stream FileStream { get; set; }
        public string FileName { get; set; }
    }
}
