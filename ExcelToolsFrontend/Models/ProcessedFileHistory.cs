namespace ExcelToolsFrontend.Models
{
    public class ProcessedFileHistory
    {
        public Guid Id { get; set; }

        public Guid FileId { get; set; }

        public string UserId { get; set; }

        public DateTime ProcessedAt { get; set; }

        public string Operation { get; set; }

    }
}
