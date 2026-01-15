using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace _2.Models
{
    public class EmployeeImportViewModel
    {
        [Required(ErrorMessage = "Vui lòng chọn file Excel")]
        [Display(Name = "File Excel")]
        public HttpPostedFileBase ExcelFile { get; set; }

        [Display(Name = "File có tiêu đề?")]
        public bool HasHeaders { get; set; } = true;

        public List<string> ImportResults { get; set; } = new List<string>();

        [Display(Name = "Thành công")]
        public int SuccessCount { get; set; }

        [Display(Name = "Lỗi")]
        public int ErrorCount { get; set; }

        [Display(Name = "Tổng cộng")]
        public int TotalCount { get { return SuccessCount + ErrorCount; } }
    }

    // Thêm class cho mapping cột Excel
    public class ExcelColumnMapping
    {
        public string ExcelColumnName { get; set; }
        public string ModelPropertyName { get; set; }
        public bool IsRequired { get; set; }
    }
}