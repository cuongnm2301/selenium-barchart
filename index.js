const excelJS = require("exceljs");
const { Builder, By, Capabilities, until } = require("selenium-webdriver");
const chrome = require("selenium-webdriver/chrome");

function extractNumbersAndPercentages(str) {
  // Regular expression to match numbers (with decimal points and signs) and percentages
  const regex = /[+-]?\d+(\.\d+)?%?/g;
  // Find all matches
  const matches = str.match(regex);

  return {
    last_price: matches[0],
    net_change_value: matches[1],
    net_change_percent: matches[2],
  };
}

function getColorForValue(value) {
  if (value && value.includes("-")) {
    return "FFFF0000"; // Red for negative values
  } else if (value && value.includes("+")) {
    return "FF00FF00"; // Green for positive values
  }
  return "FF000000";
}

function extractValues(str) {
  // Regular expression to match the pattern
  const regex = /([+-]?\d+\.\d+)\s+\(([+-]?\d+\.\d+%)\) since \d+\/\d+\/\d+/;
  const match = str.match(regex);

  if (match) {
    // match[1] is the first captured group, match[2] is the second captured group
    return { value: match[1], percentage: match[2] };
  } else {
    return { value: null, percentage: null };
  }
}

async function crawler(driver) {
  // Wait for the element containing the last price to be loaded
  let lastPriceElement = await driver.wait(
    until.elementLocated(By.className("pricechangerow")),
    10000
  );
  let last_price = await lastPriceElement.getText();
  const last_price_format = extractNumbersAndPercentages(last_price);

  let map_fill_period = [];
  let rows = await driver.findElements(By.css("table.ng-scope > tbody > tr"));
  for (let row of rows) {
    try {
      // Find the cell-period-change-title element and get its text
      let periodChangeTitle = await row
        .findElement(By.css("td.cell-period-change-title"))
        .getText();
      // Find the cell-period-change element and get its text
      let periodChange = await row
        .findElement(By.css("td.cell-period-change"))
        .getText();
      const format_period_change = periodChange.replace("\n", " ");
      const map_period_change = extractValues(format_period_change);
      map_fill_period.push({
        title: periodChangeTitle,
        ...map_period_change,
      });
    } catch (error) {
      // console.log("error: ", error);
    }
  }
  return {
    last_price_total: last_price_format,
    map_fill_period,
  };
}

const input_map = [
  {
    url: "https://www.barchart.com/futures/quotes/SBH24/performance",
    sheet: {
      last_price: "C12",
      net_change_value: "D12",
      net_change_percent: "D13",
      "5-Day_value": "E12",
      "5-Day_percent": "E13",
      "1-Month_value": "F12",
      "1-Month_percent": "F13",
      "3-Month_value": "G12",
      "3-Month_percent": "G13",
      "52-Week_value": "H12",
      "52-Week_percent": "H13",
    },
  },
  {
    url: "https://www.barchart.com/futures/quotes/SBK24/performance",
    sheet: {
      last_price: "C14",
      net_change_value: "D14",
      net_change_percent: "D15",
      "5-Day_value": "E14",
      "5-Day_percent": "E15",
      "1-Month_value": "F14",
      "1-Month_percent": "F15",
      "3-Month_value": "G14",
      "3-Month_percent": "G15",
      "52-Week_value": "H14",
      "52-Week_percent": "H15",
    },
  },
];

async function getLastPrice(input_map) {
  let chromeCapabilities = Capabilities.chrome();
  let chromeOptions = new chrome.Options();

  // Set preferences to disable redirects
  chromeOptions.setUserPreferences({
    "profile.default_content_setting_values.automatic_downloads": 2,
  });
  chromeOptions.setPageLoadStrategy("eager");
  let driver = await new Builder()
    .forBrowser("chrome")
    .setChromeOptions(chromeOptions)
    .withCapabilities(chromeCapabilities)
    .build();

  const workbook = new excelJS.Workbook();
  await workbook.xlsx.readFile("./DAILY-REPORT.xlsx");
  const worksheet = workbook.getWorksheet("Thế giới");
  try {
    for (const input of input_map) {
      await driver.get(input.url);
      const output = await crawler(driver);
      console.log("output: ", output);

      // Refactor this part to include coloring of cells
      for (let [key, cellAddress] of Object.entries(input.sheet)) {
        let cell = worksheet.getCell(cellAddress);
        let value;

        if (key in output.last_price_total) {
          value = output.last_price_total[key];
        } else {
          const periodData = output.map_fill_period.find(
            (e) => e.title === key.split("_")[0]
          );
          value = periodData
            ? key.endsWith("_value")
              ? periodData.value
              : periodData.percentage
            : null;
        }

        cell.value = value;
        cell.style = {
            ...cell.style,
            font: {...cell.font, color: {argb:  getColorForValue(value)}}
        }
      }
    }
  } finally {
    // Close the browser
    await workbook.xlsx.writeFile("./DAILY-REPORT.xlsx");
    await driver.quit();
  }
}

getLastPrice(input_map);
