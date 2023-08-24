import playwright from "playwright";
import fs from "fs";
import ExcelJS from "exceljs";

(async () => {
  const browser = await playwright.chromium.launch({
    headless: false, // set it to true to hide the browser
  });

  const page = await browser.newPage();
  await page.goto("https://finance.yahoo.com/world-indices");
  await page.waitForTimeout(5000);

  await page.locator("[name=agree]").click();
  await page.waitForTimeout(5000);

  const markets = page.locator("#market-summary li > h3");
  const stocks = [];

  for (const h3 of await markets.all()) {
    const name = await h3.getByRole("link").first().innerText();
    const value = await h3
      .locator("[data-field=regularMarketPrice]")
      .innerText();
    const change = await h3
      .locator("[data-field=regularMarketChange]")
      .innerText();
    const percent = await h3
      .locator("[data-field=regularMarketChangePercent]")
      .innerText();

    const stock = {
      name,
      value,
      change,
      percent,
    };

    stocks.push(stock);
  }

  console.log(stocks);


  // Write data to a CSV file
  const csvHeader = "Name,Value,Change,Percent\n";
  const csvRows = stocks
    .map(
      (item) => `${item.name},${item.value},${item.change},${item.percent}\n`
    )
    .join("");
  const csvContent = csvHeader + csvRows;
  fs.writeFileSync("marketSummary.csv", csvContent, "utf-8");



  //write data in excel file
  // Generate Excel file
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Data");

  // Add headers
  worksheet.addRow(["Name", "value", "change", "percent"]);

  // Add data
  stocks.forEach((item) => {
    worksheet.addRow([item.name, item.value, item.change, item.percent]);
  });

  // Save the workbook to a file
  const excelFilePath = "marketSummary.xlsx";
  await workbook.xlsx.writeFile(excelFilePath);

  await browser.close();
})();
