using Microsoft.AspNetCore.Components;
using InvoiceApplication;
using Aspose.Words;
using Aspose.Cells;
using System.IO;

namespace InvoiceApplication.Pages
{
	public partial class GenerateInvoice : ComponentBase
	{
		[Inject] private IWebHostEnvironment Environment { get; set; } = default!;

		private Invoice invoice = new() { Date = DateTime.Now };

		protected override void OnInitialized()
		{
			AddItem();
		}

		private void AddItem()
		{
			invoice.Items.Add(new InvoiceItem { Quantity = 1, UnitPrice = 0 });
			CalculateTotal();
		}

		private void RemoveItem(InvoiceItem itemToRemove)
		{
			invoice.Items.Remove(itemToRemove);
			CalculateTotal();
		}

		private void CalculateTotal()
		{
			invoice.TotalAmount = invoice.Items.Sum(item => item.Quantity * item.UnitPrice);
		}

		private void OnQuantityChanged(InvoiceItem item, string? value)
		{
			if (int.TryParse(value, out int qty))
			{
				item.Quantity = qty;
				CalculateTotal();
			}
		}

		private void OnUnitPriceChanged(InvoiceItem item, string? value)
		{
			if (decimal.TryParse(value, out decimal price))
			{
				item.UnitPrice = price;
				CalculateTotal();
			}
		}

		private void GeneratePdf() => GeneratePdfFile(GetFilePath("pdf"));
		private void GenerateWord() => GenerateWordFile(GetFilePath("docx"));
		private void GenerateExcel() => GenerateExcelFile(GetFilePath("xlsx"));

		private string GetFilePath(string ext) =>
			$"wwwroot/invoices/invoice_{DateTime.Now.Ticks}.{ext}";

		private void EnsureDirectory()
		{
			var directory = Path.Combine(Environment.ContentRootPath, "wwwroot", "invoices");
			if (!Directory.Exists(directory))
				Directory.CreateDirectory(directory);
		}

		private void GenerateWordFile(string fileName)
		{
			EnsureDirectory();
			var doc = new Document();
			var builder = new DocumentBuilder(doc);

			builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
			builder.Font.Size = 18;
			builder.Writeln("INVOICE");

			builder.InsertParagraph();
			builder.Font.Size = 12;
			builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
			builder.Writeln($"Date: {invoice.Date.ToShortDateString()}");
			builder.Writeln($"Client: {invoice.ClientName}");
			builder.Writeln($"Address: {invoice.ClientAddress}");

			builder.InsertParagraph();
			builder.StartTable();
			builder.InsertCell(); builder.Font.Bold = true; builder.Write("Product");
			builder.InsertCell(); builder.Write("Quantity");
			builder.InsertCell(); builder.Write("Unit Price");
			builder.InsertCell(); builder.Write("Total");
			builder.EndRow();

			foreach (var item in invoice.Items)
			{
				builder.InsertCell(); builder.Font.Bold = false; builder.Write(item.ProductName);
				builder.InsertCell(); builder.Write(item.Quantity.ToString());
				builder.InsertCell(); builder.Write($"{item.UnitPrice:C}");
				builder.InsertCell(); builder.Write($"{item.Total:C}");
				builder.EndRow();
			}

			builder.InsertCell(); builder.InsertCell(); builder.InsertCell(); builder.Font.Bold = true;
			builder.Write($"Total: {invoice.TotalAmount:C}");
			builder.EndRow();
			builder.EndTable();

			builder.InsertParagraph();
			builder.Writeln("Thank you for your business!");

			doc.Save(Path.Combine(Environment.ContentRootPath, fileName));
		}

		private void GeneratePdfFile(string fileName)
		{
			GenerateWordFile(fileName.Replace(".pdf", ".docx"));
			var doc = new Document(Path.Combine(Environment.ContentRootPath, fileName.Replace(".pdf", ".docx")));
			doc.Save(Path.Combine(Environment.ContentRootPath, fileName), Aspose.Words.SaveFormat.Pdf);
		}

		private void GenerateExcelFile(string fileName)
		{
			EnsureDirectory();
			var workbook = new Workbook();
			var sheet = workbook.Worksheets[0];

			int row = 0;
			sheet.Cells[row++, 0].PutValue("INVOICE");
			sheet.Cells[row++, 0].PutValue($"Client: {invoice.ClientName}");
			sheet.Cells[row++, 0].PutValue($"Address: {invoice.ClientAddress}");
			sheet.Cells[row++, 0].PutValue($"Date: {invoice.Date.ToShortDateString()}");

			sheet.Cells[row, 0].PutValue("Product");
			sheet.Cells[row, 1].PutValue("Quantity");
			sheet.Cells[row, 2].PutValue("Unit Price");
			sheet.Cells[row, 3].PutValue("Total");
			row++;

			foreach (var item in invoice.Items)
			{
				sheet.Cells[row, 0].PutValue(item.ProductName);
				sheet.Cells[row, 1].PutValue(item.Quantity);
				sheet.Cells[row, 2].PutValue(item.UnitPrice);
				sheet.Cells[row, 3].PutValue(item.Total);
				row++;
			}

			sheet.Cells[row, 2].PutValue("Total:");
			sheet.Cells[row, 3].PutValue(invoice.TotalAmount);

			workbook.Save(Path.Combine(Environment.ContentRootPath, fileName));
		}
	}
}
