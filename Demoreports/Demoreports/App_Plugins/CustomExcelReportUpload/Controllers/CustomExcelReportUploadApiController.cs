using System;
using System.IO;
using System.Web;
using System.Web.Mvc;
using System.Data.OleDb;
using Umbraco.Core.Models;
using Umbraco.Web.Mvc;
using System.Linq;
using System.Globalization;

namespace CustomExcelReportUpload.Controllers
{
	public class CustomExcelReportUploadApiController : SurfaceController
	{
		[HttpPost]
		public ActionResult Upload(HttpPostedFileBase file)
		{
			try
			{
				if (file != null && file.ContentLength > 0)
				{
					var fileExtension = Path.GetExtension(file.FileName);

					if (fileExtension.Equals(".xls") || fileExtension.Equals(".xlsx"))
					{
						var tempFilePath = Path.Combine(Path.GetTempPath(), file.FileName);
						file.SaveAs(tempFilePath);

						var connectionString = GetConnectionString(tempFilePath);
						using (var connection = new OleDbConnection(connectionString))
						{
							connection.Open();
							var command = new OleDbCommand("SELECT * FROM [Sheet1$]", connection);
							using (var reader = command.ExecuteReader())
							{
								while (reader.Read())
								{
									var auditSiteName = reader.GetValue(0).ToString();
									var siteAddress = reader.GetValue(1).ToString();
									var standardName = reader.GetValue(2).ToString();
									var certificationScope = reader.GetValue(3).ToString();
									var certificateReportNumber = reader.GetValue(4).ToString();
									var expiryDateString = reader.GetValue(5).ToString();

									DateTime expiryDate;
									bool expire = false;
									if (!DateTime.TryParseExact(expiryDateString, "dd MMMM yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out expiryDate))
									{
										expire = true;
									}


									if (!ReportItemExists(standardName, certificationScope) && !expire)
									{
										CreateReportsItem(auditSiteName, siteAddress, standardName, certificationScope, certificateReportNumber, expiryDateString);
									}

								}
							}
						}

						System.IO.File.Delete(tempFilePath);

						return Json(new { success = true, message = "File processed +" });
					}
					else
					{
						return Json(new { success = false, message = "Invalid file format. Please upload an Excel file." });
					}
				}
				return Json(new { success = false, message = "No file received" });
			}
			catch (Exception ex)
			{
				return Json(new { success = false, message = ex.Message });
			}
		}

		private string GetConnectionString(string filePath)
		{
			if (Path.GetExtension(filePath).Equals(".xls"))
			{
				return $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={filePath};Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1;\"";
			}
			else if (Path.GetExtension(filePath).Equals(".xlsx"))
			{
				return $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1;\"";
			}
			else
			{
				throw new Exception("Invalid Excel file format");
			}
		}

		private bool ReportItemExists(string standardName, string certificationScope)
		{
			var rootContent = Services.ContentService.GetRootContent().FirstOrDefault(x => x.ContentType.Alias == "reports");
			if (rootContent != null)
			{
				var contentTypeService = Services.ContentTypeService;
				var reportsItemType = contentTypeService.GetAll().FirstOrDefault(x => x.Alias == "reportsItem");
				if (reportsItemType != null)
				{
					var childNodes = Services.ContentService.GetPagedChildren(rootContent.Id, 0, int.MaxValue, out long totalRecords)
						.Where(x => x.ContentType.Id == reportsItemType.Id);

					var existingItem = childNodes.FirstOrDefault(x => x.GetValue<string>("standardName") == standardName && x.GetValue<string>("certificationScope") == certificationScope);

					return existingItem != null;
				}
			}
			return false;
		}

		private void CreateReportsItem(string auditSiteName, string siteAddress, string standardName, string certificationScope, string certificateReportNumber, string expiryDate)
		{
			var rootContent = Services.ContentService.GetRootContent().FirstOrDefault(x => x.ContentType.Alias == "reports");
			if (rootContent != null)
			{
				var reportsItem = Services.ContentService.Create("ReportsItem", rootContent.Id, "reportsItem");
				reportsItem.SetValue("auditSiteName", auditSiteName);
				reportsItem.SetValue("siteAddress", siteAddress);
				reportsItem.SetValue("standardName", standardName);
				reportsItem.SetValue("certificationScope", certificationScope);
				reportsItem.SetValue("certificateReportNumber", certificateReportNumber);
				reportsItem.SetValue("expiryDate", expiryDate);

				Services.ContentService.SaveAndPublish(reportsItem);
			}
		}
	}
}
