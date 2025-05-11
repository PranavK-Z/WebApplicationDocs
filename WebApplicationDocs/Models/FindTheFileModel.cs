using System.ComponentModel.DataAnnotations;

namespace WebApplicationDocs.Models
{
    public class FindTheFileModel
    {

        [Required]
        public string SourcePath { get; set; }

        [Required]
        public string ClientId { get; set; }

        [Required]
        [StringLength(3, MinimumLength = 3, ErrorMessage = "Document type must be exactly 3 characters.")]
        public string DocumentType { get; set; }

        [Required]
        public string DestPath { get; set; }

        [Required]
        [Range(1, int.MaxValue, ErrorMessage = "Enter a valid number of Payment Files.")]
        public int PaymentFileNum { get; set; }


    }
}

