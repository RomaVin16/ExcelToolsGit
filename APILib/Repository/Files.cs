using System.ComponentModel.DataAnnotations.Schema;

namespace APILib.Repository
{
    public class Files
    {
        public Guid Id { get; set; }
        public string FileName { get; set; }

		public DateTime SaveDate { get; set; }
        public FileState State { get; set; }
        public long SizeBytes { get; set; }

        [Column("operation")]
        public string Operation { get; set; }

		public enum FileState
        {
            Active,
            Removed
        }
    }
}
