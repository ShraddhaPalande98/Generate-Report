using DataTables.Data;
using DataTables.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using PdfSharpCore;
using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;
using System.Diagnostics;

namespace DataTables.Controllers
{
    public class HomeController : Controller
    {
        private readonly ApplicationDbContext _context;

        public HomeController(ApplicationDbContext context)
        {
            _context = context;
        }

        [HttpGet]
        public async Task<IActionResult> Index(string sortOrder, string searchString, int page = 1)
        {
            ViewData["NameSortParm"] = String.IsNullOrEmpty(sortOrder) ? "name_desc" : "";
            ViewData["CategorySortParm"] = String.IsNullOrEmpty(sortOrder) ? "category_desc" : "";
            ViewData["PriceSortParm"] = sortOrder == "price_asc" ? "price_desc" : "price_asc";

            var product = from p in _context.Sells 
                          //orderby p.Id
                          select p;

            if (!string.IsNullOrEmpty(searchString))
            {
                product = product.Where(p => p.ProductName.Contains(searchString));
            }

            switch (sortOrder)
            {
                case "name_desc":
                    product = product.OrderByDescending(p => p.ProductName);
                    break;
                case "category_desc":
                    product = product.OrderByDescending(p => p.Price);
                    break;
                case "price_asc":
                    product = product.OrderBy(p => p.Price);
                    break;
                case "price_desc":
                    product = product.OrderByDescending(p => p.Price);
                    break;
                default:
                    product = product.OrderBy(p => p.Id);
                    break;
            }

            int pageSize = 10;
            var totalItems = await product.CountAsync();
            var totalPages = (int)Math.Ceiling((double)totalItems / pageSize);
            var pagedProducts = await product.Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();

            ViewData["CurrentPage"] = page;
            ViewData["TotalPages"] = totalPages;
            ViewData["TotalItems"] = totalItems;
            ViewData["PageSize"] = pageSize;
            ViewData["SearchString"] = searchString;

            return View(pagedProducts);

        }

        public async Task<IActionResult> ExportToExcel(string searchString, int page, int pageSize)
        {

            var products = from p in _context.Sells
                           select p;

            if (!string.IsNullOrEmpty(searchString))
            {
                products = products.Where(p => p.ProductName.Contains(searchString));
            }

            var productList = await products.Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Products");
                worksheet.Cells[1, 1].Value = "ID";
                worksheet.Cells[1, 2].Value = "Product Name";
                worksheet.Cells[1, 3].Value = "Description";
                worksheet.Cells[1, 4].Value = "Category";
                worksheet.Cells[1, 5].Value = "Price";
                worksheet.Cells[1, 6].Value = "Quantity";

                for (int i = 0; i < productList.Count; i++)
                {
                    var product = productList[i];
                    worksheet.Cells[i + 2, 1].Value = product.Id;
                    worksheet.Cells[i + 2, 2].Value = product.ProductName;
                    worksheet.Cells[i + 2, 3].Value = product.ProductDescription;
                    worksheet.Cells[i + 2, 4].Value = product.ProductCategory;
                    worksheet.Cells[i + 2, 5].Value = product.Price;
                    worksheet.Cells[i + 2, 6].Value = product.Quantity;
                }

                var stream = new MemoryStream(package.GetAsByteArray());
                return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ProductList.xlsx");
            }
        }


        public async Task<IActionResult> ExportToPDF(string searchString, int page = 1, int pageSize = 10)
        {
            var productsQuery = from p in _context.Sells
                                select p;

            if (!string.IsNullOrEmpty(searchString))
            {
                productsQuery = productsQuery.Where(p => p.ProductName.Contains(searchString));
            }

            var pagedProducts = await productsQuery.Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();

            var stream = new MemoryStream();
            var document = new PdfDocument();

            var pdfPage = document.AddPage();
            var gfx = XGraphics.FromPdfPage(pdfPage);
            var font = new XFont("Arial", 12);

            gfx.DrawString("Product List", new XFont("Arial", 18, XFontStyle.Bold), XBrushes.Black, new XRect(0, 40, pdfPage.Width, 50), XStringFormats.Center);

            int yPosition = 80;
            gfx.DrawString("ID", font, XBrushes.Black, 40, yPosition);
            gfx.DrawString("Product Name", font, XBrushes.Black, 100, yPosition);
            gfx.DrawString("Description", font, XBrushes.Black, 250, yPosition);
            gfx.DrawString("Category", font, XBrushes.Black, 400, yPosition);
            gfx.DrawString("Price", font, XBrushes.Black, 500, yPosition);
            gfx.DrawString("Quantity", font, XBrushes.Black, 600, yPosition);

            yPosition += 30; 

            foreach (var product in pagedProducts)
            {
                gfx.DrawString(product.Id.ToString(), font, XBrushes.Black, 40, yPosition);
                gfx.DrawString(product.ProductName, font, XBrushes.Black, 100, yPosition);
                gfx.DrawString(product.ProductDescription, font, XBrushes.Black, 250, yPosition);
                gfx.DrawString(product.ProductCategory, font, XBrushes.Black, 400, yPosition);
                gfx.DrawString(product.Price.ToString("C"), font, XBrushes.Black, 500, yPosition);
                gfx.DrawString(product.Quantity.ToString(), font, XBrushes.Black, 600, yPosition);

                yPosition += 20;
            }

            document.Save(stream);
            stream.Position = 0;
            return File(stream, "application/pdf", "ProductList.pdf");
        }

    }
}
