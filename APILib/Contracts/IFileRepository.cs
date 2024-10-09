using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using APILib.Repository;

namespace APILib.Contracts
{
    public interface IFileRepository
    {
        Task Create(Stream stream, Guid fileId, string fileName);
        string GetFileName(Guid fileId);
        void Update();
    }
}
