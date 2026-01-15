using _2.Models;
using ExcelDataReader;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Validation;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using System.Drawing;
namespace _2.Controllers
{
    public class EmployeesController : Controller
    {
        private QuanLyNhanSuEntities db = new QuanLyNhanSuEntities();

        // GET: Employees
        public ActionResult Index()
        {
            var employees = db.Employees.Include(e => e.Company);
            return View(employees.ToList());
        }

        // GET: Employees/Details/5
        public ActionResult Details(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Employee employee = db.Employees.Find(id);
            if (employee == null)
            {
                return HttpNotFound();
            }
            return View(employee);
        }

        // GET: Employees/Create
        public ActionResult Create()
        {
            ViewBag.company_id = new SelectList(db.Companies, "company_id", "company_name");
            return View();
        }

        // POST: Employees/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "emp_id,emp_name,emp_sex,emp_position,job_position,supervisor_id,company_id")] Employee employee)
        {
            try
            {
                if (ModelState.IsValid)
                {
                    // Kiểm tra emp_id đã tồn tại chưa
                    var existingEmployee = db.Employees.FirstOrDefault(e => e.emp_id == employee.emp_id);
                    if (existingEmployee != null)
                    {
                        ModelState.AddModelError("emp_id", "Mã nhân viên đã tồn tại trong hệ thống.");

                        ViewBag.company_id = new SelectList(db.Companies, "company_id", "company_name", employee.company_id);
                        return Json(new
                        {
                            success = false,
                            errors = new Dictionary<string, string> {
                        { "emp_id", "Mã nhân viên đã tồn tại trong hệ thống." }
                    }
                        });
                    }

                    db.Employees.Add(employee);
                    db.SaveChanges();
                    return Json(new { success = true });
                }

                // Nếu ModelState không valid, trả về tất cả lỗi
                var errors = new Dictionary<string, string>();
                foreach (var state in ModelState)
                {
                    if (state.Value.Errors.Any())
                    {
                        errors[state.Key] = string.Join(", ", state.Value.Errors.Select(e => e.ErrorMessage));
                    }
                }

                return Json(new { success = false, errors = errors });
            }
            catch (Exception ex)
            {
                // Log lỗi nếu cần
                return Json(new
                {
                    success = false,
                    errors = new Dictionary<string, string> {
                { "", "Đã xảy ra lỗi: " + ex.Message }
            }
                });
            }
        }
        public JsonResult CheckEmployeeId(string emp_id)
        {
            if (string.IsNullOrEmpty(emp_id))
                return Json(true, JsonRequestBehavior.AllowGet);

            var exists = db.Employees.Any(e => e.emp_id == emp_id);
            return Json(!exists, JsonRequestBehavior.AllowGet);
        }

        // GET: Employees/Edit/5
        public ActionResult Edit(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Employee employee = db.Employees.Find(id);
            if (employee == null)
            {
                return HttpNotFound();
            }
            ViewBag.company_id = new SelectList(db.Companies, "company_id", "company_name", employee.company_id);
            return View(employee);
        }

        // POST: Employees/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "emp_id,emp_name,emp_sex,emp_position,job_position,supervisor_id,company_id")] Employee employee)
        {
            if (ModelState.IsValid)
            {
                db.Entry(employee).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            ViewBag.company_id = new SelectList(db.Companies, "company_id", "company_name", employee.company_id);
            return View(employee);
        }

        // GET: Employees/Delete/5
        public ActionResult Delete(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Employee employee = db.Employees.Find(id);
            if (employee == null)
            {
                return HttpNotFound();
            }
            return View(employee);
        }

        // POST: Employees/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(string id)
        {
            Employee employee = db.Employees.Find(id);
            db.Employees.Remove(employee);
            db.SaveChanges();
            return RedirectToAction("OrganizationalChart");
        }

        public ActionResult OrganizationalChart()
        {
            // Lấy tất cả nhân viên từ database
            var employees = db.Employees.ToList();

            // Truyền dữ liệu sang view
            ViewBag.TotalEmployees = employees.Count;
            ViewBag.TotalMale = employees.Count(e => e.emp_sex == "Nam");
            ViewBag.TotalFemale = employees.Count(e => e.emp_sex == "Nữ");

            return View(employees);
        }
        public JsonResult GetEmployeesJson()
        {
            try
            {
                var employees = db.Employees
                    .Select(e => new
                    {
                        e.emp_id,
                        e.emp_name,
                        e.emp_sex,
                        e.emp_position,
                        e.job_position,
                        e.supervisor_id
                    })
                    .ToList();

                return Json(employees, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                // Log error
                System.Diagnostics.Debug.WriteLine($"Error: {ex.Message}");
                return Json(new { error = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        public ActionResult Import()
        {
            return View(new EmployeeImportViewModel());
        }

        // POST: Import từ Excel
        [HttpPost]
        public ActionResult ImportExcel(HttpPostedFileBase file, bool hasHeaders = true, bool skipDuplicates = true)
        {
            try
            {
                if (file == null || file.ContentLength == 0)
                {
                    return Json(new { success = false, message = "Vui lòng chọn file Excel" });
                }

                var fileExt = Path.GetExtension(file.FileName).ToLower();
                if (fileExt != ".xlsx" && fileExt != ".xls")
                {
                    return Json(new { success = false, message = "Chỉ hỗ trợ file Excel (.xlsx, .xls)" });
                }

                var employees = ReadEmployeesFromExcel(file.InputStream, hasHeaders);
                var results = new List<object>();
                int successCount = 0;
                int errorCount = 0;

                foreach (var emp in employees)
                {
                    try
                    {
                        // Kiểm tra dữ liệu bắt buộc
                        if (string.IsNullOrEmpty(emp.emp_id) || string.IsNullOrEmpty(emp.emp_name))
                        {
                            results.Add(new
                            {
                                status = "error",
                                message = $"Dòng {successCount + errorCount + 1}: Thiếu ID hoặc tên"
                            });
                            errorCount++;
                            continue;
                        }

                        // Kiểm tra trùng
                        var existing = db.Employees.FirstOrDefault(e => e.emp_id == emp.emp_id);
                        if (existing != null)
                        {
                            if (skipDuplicates)
                            {
                                results.Add(new
                                {
                                    status = "warning",
                                    message = $"Bỏ qua: {emp.emp_id} - Mã đã tồn tại"
                                });
                                continue;
                            }
                            else
                            {
                                results.Add(new
                                {
                                    status = "error",
                                    message = $"Lỗi: {emp.emp_id} - Mã đã tồn tại"
                                });
                                errorCount++;
                                continue;
                            }
                        }

                        // Thiết lập giá trị mặc định
                        if (!emp.company_id.HasValue)
                        {
                            emp.company_id = 1;
                        }

                        if (string.IsNullOrEmpty(emp.emp_sex))
                        {
                            emp.emp_sex = "Nam";
                        }

                        // Thêm vào database
                        db.Employees.Add(emp);
                        db.SaveChanges();

                        results.Add(new
                        {
                            status = "success",
                            message = $"Thành công: {emp.emp_id} - {emp.emp_name}"
                        });
                        successCount++;
                    }
                    catch (Exception ex)
                    {
                        results.Add(new
                        {
                            status = "error",
                            message = $"Lỗi: {emp.emp_id} - {ex.Message}"
                        });
                        errorCount++;
                    }
                }

                return Json(new
                {
                    success = true,
                    results = results,
                    successCount = successCount,
                    errorCount = errorCount,
                    total = employees.Count
                });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = $"Lỗi hệ thống: {ex.Message}" });
            }
        }

        private List<Employee> ReadEmployeesFromExcel(Stream stream, bool hasHeaders)
        {
            var employees = new List<Employee>();

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (data) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = hasHeaders
                    }
                });

                var dataTable = result.Tables[0];

                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    var row = dataTable.Rows[i];

                    // Bỏ qua dòng trống
                    if (IsEmptyRow(row))
                        continue;

                    var employee = new Employee
                    {
                        emp_id = GetCellValue(row[0]),
                        emp_name = GetCellValue(row[1]),
                        emp_sex = GetCellValue(row[2]),
                        emp_position = GetCellValue(row[3]),
                        job_position = GetCellValue(row[4]),
                        supervisor_id = GetCellValue(row[5]),
                        company_id = GetIntValue(row[6])
                    };

                    employees.Add(employee);
                }
            }

            return employees;
        }

        private string GetCellValue(object cell)
        {
            if (cell == null || cell == DBNull.Value)
                return string.Empty;

            return cell.ToString().Trim();
        }

        private bool IsEmptyRow(System.Data.DataRow row)
        {
            for (int i = 0; i < row.ItemArray.Length; i++)
            {
                if (!string.IsNullOrEmpty(GetCellValue(row[i])))
                    return false;
            }
            return true;
        }

        // Helper methods
        private string GetStringValue(object cellValue)
        {
            if (cellValue == null || cellValue == DBNull.Value)
                return null;

            return cellValue.ToString().Trim();
        }

        private int? GetIntValue(object cellValue, int? defaultValue = null)
        {
            if (cellValue == null || cellValue == DBNull.Value)
                return defaultValue;

            string stringValue = cellValue.ToString().Trim();

            if (string.IsNullOrEmpty(stringValue))
                return defaultValue;

            // Thử parse sang int
            if (int.TryParse(stringValue, out int intValue))
                return intValue;

            // Nếu không parse được, trả về giá trị mặc định
            return defaultValue;
        }



        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
