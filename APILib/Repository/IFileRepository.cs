using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace APILib.Repository
{
    public interface IFileRepository
    {
        Task Create(Files files);
    }
}
