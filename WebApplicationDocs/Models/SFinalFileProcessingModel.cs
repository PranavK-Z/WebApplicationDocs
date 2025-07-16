using System.ComponentModel.DataAnnotations;

namespace WebApplicationDocs.Models
{
    public class SFinalFileProcessingModel
    {
        [Required]
        public string SourcePath { get; set; }

        [Required]
        public string ClientId { get; set; }
        [Required]
        public string RecipientType { get; set; }

        [Required]
        [StringLength(3, MinimumLength = 3, ErrorMessage = "Document type must be exactly 3 characters.")]
        public string DocumentType { get; set; }

        [StringLength(2, MinimumLength = 2, ErrorMessage = "Replacement suffix must be exactly 2 characters.")]
        public string ReplacementSuffix { get; set; }

        [Required]
        public string DestPath { get; set; }

        [Required]
        [Range(1, int.MaxValue, ErrorMessage = "Enter a valid number of Payment Files.")]
        public int PaymentFileNum { get; set; }

        public IFormFile ExcelFile { get; set; }
    }

    public class SFileProcessingModelNewWithoutExcelAndSuffix : SFinalFileProcessingModel
    {
        [Required(AllowEmptyStrings = true)]
        public new string ReplacementSuffix { get; set; } = null;

        [Required(AllowEmptyStrings = true)]
        public new IFormFile ExcelFile { get; set; } = null;
    }

}

