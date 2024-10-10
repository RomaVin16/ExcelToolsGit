using Microsoft.AspNetCore.Mvc;
using System.ComponentModel.DataAnnotations;

namespace API.Models
{
    public class FileUploadRequest
    {
        [Required]
        [FromForm]
        public IFormFile[] Files { get; set; }
    }
}
