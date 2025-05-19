using System.ComponentModel.DataAnnotations;

namespace ExcelToolsFrontend.Models
{
    public class OtpSendModel
    {
        [Required(ErrorMessage = "Введите номер телефона")]
        [Phone(ErrorMessage = "Неверный формат номера")]
        public string PhoneNumber { get; set; } = "";
    }
}
