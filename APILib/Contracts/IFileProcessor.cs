using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using API.Models;

namespace APILib.Contracts
{
    public interface IFileProcessor
    {
        Task<Guid> UploadAsync(string fileName, Stream stream);
        FileResult Download(Guid fileId);
    }
}
