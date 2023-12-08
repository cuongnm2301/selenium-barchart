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
  {
    url: "https://www.barchart.com/futures/quotes/SBN24/overview",
    sheet: {
      last_price: "C16",
      net_change_value: "D16",
      net_change_percent: "D17",
      "5-Day_value": "E16",
      "5-Day_percent": "E17",
      "1-Month_value": "F16",
      "1-Month_percent": "F17",
      "3-Month_value": "G16",
      "3-Month_percent": "G17",
      "52-Week_value": "H16",
      "52-Week_percent": "H17"
    },
  },
  {
    "url": "https://www.barchart.com/futures/quotes/SBV24/overview",
    "sheet": {
      "last_price": "C18",
      "net_change_value": "D18",
      "net_change_percent": "D19",
      "5-Day_value": "E18",
      "5-Day_percent": "E19",
      "1-Month_value": "F18",
      "1-Month_percent": "F19",
      "3-Month_value": "G18",
      "3-Month_percent": "G19",
      "52-Week_value": "H18",
      "52-Week_percent": "H19"
    }
  },
  {
    "url": "https://www.barchart.com/futures/quotes/SWH24/overview",
    "sheet": {
      "last_price": "C21",
      "net_change_value": "D21",
      "net_change_percent": "D22",
      "5-Day_value": "E21",
      "5-Day_percent": "E22",
      "1-Month_value": "F21",
      "1-Month_percent": "F22",
      "3-Month_value": "G21",
      "3-Month_percent": "G22",
      "52-Week_value": "H21",
      "52-Week_percent": "H22"
    }
  },
  {
    "url": "https://www.barchart.com/futures/quotes/SWk24/overview",
    "sheet": {
      "last_price": "C23",
      "net_change_value": "D23",
      "net_change_percent": "D24",
      "5-Day_value": "E23",
      "5-Day_percent": "E24",
      "1-Month_value": "F23",
      "1-Month_percent": "F24",
      "3-Month_value": "G23",
      "3-Month_percent": "G24",
      "52-Week_value": "H23",
      "52-Week_percent": "H24"
    }
  },
  {
    "url": "https://www.barchart.com/futures/quotes/SWq24/overview",
    "sheet": {
      "last_price": "C25",
      "net_change_value": "D25",
      "net_change_percent": "D26",
      "5-Day_value": "E25",
      "5-Day_percent": "E26",
      "1-Month_value": "F25",
      "1-Month_percent": "F26",
      "3-Month_value": "G25",
      "3-Month_percent": "G26",
      "52-Week_value": "H25",
      "52-Week_percent": "H26"
    }
  },
  {
    "url": "https://www.barchart.com/futures/quotes/SWV24/overview",
    "sheet": {
      "last_price": "C27",
      "net_change_value": "D27",
      "net_change_percent": "D28",
      "5-Day_value": "E27",
      "5-Day_percent": "E28",
      "1-Month_value": "F27",
      "1-Month_percent": "F28",
      "3-Month_value": "G27",
      "3-Month_percent": "G28",
      "52-Week_value": "H27",
      "52-Week_percent": "H28"
    }
  },
  {
    "url": "https://www.barchart.com/futures/quotes/SWZ24/overview",
    "sheet": {
      "last_price": "C29",
      "net_change_value": "D29",
      "net_change_percent": "D30",
      "5-Day_value": "E29",
      "5-Day_percent": "E30",
      "1-Month_value": "F29",
      "1-Month_percent": "F30",
      "3-Month_value": "G29",
      "3-Month_percent": "G30",
      "52-Week_value": "H29",
      "52-Week_percent": "H30"
    }
  },
  {
    "url": "https://www.barchart.com/futures/quotes/GCY00/overview",
    "sheet": {
      "last_price": "C32",
      "net_change_value": "D32",
      "net_change_percent": "D33",
      "5-Day_value": "E32",
      "5-Day_percent": "E33",
      "1-Month_value": "F32",
      "1-Month_percent": "F33",
      "3-Month_value": "G32",
      "3-Month_percent": "G33",
      "52-Week_value": "H32",
      "52-Week_percent": "H33"
    }
  },
  {
      "url": "https://www.barchart.com/futures/quotes/CLF24/overview",
      "sheet": {
        "last_price": "C35",
        "net_change_value": "D35",
        "net_change_percent": "D36",
        "5-Day_value": "E35",
        "5-Day_percent": "E36",
        "1-Month_value": "F35",
        "1-Month_percent": "F36",
        "3-Month_value": "G35",
        "3-Month_percent": "G36",
        "52-Week_value": "H35",
        "52-Week_percent": "H36"
      }
    },
    {
      "url": "https://www.barchart.com/futures/quotes/QAG24/overview",
      "sheet": {
        "last_price": "C37",
        "net_change_value": "D37",
        "net_change_percent": "D38",
        "5-Day_value": "E37",
        "5-Day_percent": "E38",
        "1-Month_value": "F37",
        "1-Month_percent": "F38",
        "3-Month_value": "G37",
        "3-Month_percent": "G38",
        "52-Week_value": "H37",
        "52-Week_percent": "H38"
      }
    },
    {
      "url": "https://www.barchart.com/futures/quotes/RBF24/overview",
      "sheet": {
        "last_price": "C39",
        "net_change_value": "D39",
        "net_change_percent": "D40",
        "5-Day_value": "E39",
        "5-Day_percent": "E40",
        "1-Month_value": "F39",
        "1-Month_percent": "F40",
        "3-Month_value": "G39",
        "3-Month_percent": "G40",
        "52-Week_value": "H39",
        "52-Week_percent": "H40"
      }
    },
    {
      "url": "https://www.barchart.com/futures/quotes/ZKY00/overview",
      "sheet": {
        "last_price": "C41",
        "net_change_value": "D41",
        "net_change_percent": "D42",
        "5-Day_value": "E41",
        "5-Day_percent": "E42",
        "1-Month_value": "F41",
        "1-Month_percent": "F42",
        "3-Month_value": "G41",
        "3-Month_percent": "G42",
        "52-Week_value": "H41",
        "52-Week_percent": "H42"
      }
    },
    {
        "url": "https://www.barchart.com/stocks/quotes/$DXY/overview",
        "sheet": {
          "last_price": "C44",
          "net_change_value": "D44",
          "net_change_percent": "D45",
          "5-Day_value": "E44",
          "5-Day_percent": "E45",
          "1-Month_value": "F44",
          "1-Month_percent": "F45",
          "3-Month_value": "G44",
          "3-Month_percent": "G45",
          "52-Week_value": "H44",
          "52-Week_percent": "H45"
        }
    },
    {
        "url": "https://www.barchart.com/forex/quotes/%5EUSDAUD/overview",
        "sheet": {
          "last_price": "C46",
          "net_change_value": "D46",
          "net_change_percent": "D47",
          "5-Day_value": "E46",
          "5-Day_percent": "E47",
          "1-Month_value": "F46",
          "1-Month_percent": "F47",
          "3-Month_value": "G46",
          "3-Month_percent": "G47",
          "52-Week_value": "H46",
          "52-Week_percent": "H47"
        }
    },
    {
        "url": "https://www.barchart.com/forex/quotes/%5EUSDBRL/overview",
        "sheet": {
          "last_price": "C48",
          "net_change_value": "D48",
          "net_change_percent": "D49",
          "5-Day_value": "E48",
          "5-Day_percent": "E49",
          "1-Month_value": "F48",
          "1-Month_percent": "F49",
          "3-Month_value": "G48",
          "3-Month_percent": "G49",
          "52-Week_value": "H48",
          "52-Week_percent": "H49"
        }
    },
    {
        "url": "https://www.barchart.com/forex/quotes/%5EUSDEUR/overview",
        "sheet": {
          "last_price": "C50",
          "net_change_value": "D50",
          "net_change_percent": "D51",
          "5-Day_value": "E50",
          "5-Day_percent": "E51",
          "1-Month_value": "F50",
          "1-Month_percent": "F51",
          "3-Month_value": "G50",
          "3-Month_percent": "G51",
          "52-Week_value": "H50",
          "52-Week_percent": "H51"
        }
    },
    {
        "url": "https://www.barchart.com/forex/quotes/%5EUSDJPY/overview",
        "sheet": {
          "last_price": "C52",
          "net_change_value": "D52",
          "net_change_percent": "D53",
          "5-Day_value": "E52",
          "5-Day_percent": "E53",
          "1-Month_value": "F52",
          "1-Month_percent": "F53",
          "3-Month_value": "G52",
          "3-Month_percent": "G53",
          "52-Week_value": "H52",
          "52-Week_percent": "H53"
        }
    },
    {
        "url": "https://www.barchart.com/forex/quotes/%5EUSDINR/overview",
        "sheet": {
          "last_price": "C54",
          "net_change_value": "D54",
          "net_change_percent": "D55",
          "5-Day_value": "E54",
          "5-Day_percent": "E55",
          "1-Month_value": "F54",
          "1-Month_percent": "F55",
          "3-Month_value": "G54",
          "3-Month_percent": "G55",
          "52-Week_value": "H54",
          "52-Week_percent": "H55"
        }
    },
    {
        "url": "https://www.barchart.com/forex/quotes/%5EUSDCNY/overview",
        "sheet": {
          "last_price": "C56",
          "net_change_value": "D56",
          "net_change_percent": "D57",
          "5-Day_value": "E56",
          "5-Day_percent": "E57",
          "1-Month_value": "F56",
          "1-Month_percent": "F57",
          "3-Month_value": "G56",
          "3-Month_percent": "G57",
          "52-Week_value": "H56",
          "52-Week_percent": "H57"
        }
    },
    {
        "url": "https://www.barchart.com/forex/quotes/%5EUSDTHB/overview",
        "sheet": {
          "last_price": "C58",
          "net_change_value": "D58",
          "net_change_percent": "D59",
          "5-Day_value": "E58",
          "5-Day_percent": "E59",
          "1-Month_value": "F58",
          "1-Month_percent": "F59",
          "3-Month_value": "G58",
          "3-Month_percent": "G59",
          "52-Week_value": "H58",
          "52-Week_percent": "H59"
        }
    },
    {
        "url": "https://www.barchart.com/forex/quotes/%5EUSDVND/overview",
        "sheet": {
          "last_price": "C60",
          "net_change_value": "D60",
          "net_change_percent": "D61",
          "5-Day_value": "E60",
          "5-Day_percent": "E61",
          "1-Month_value": "F60",
          "1-Month_percent": "F61",
          "3-Month_value": "G60",
          "3-Month_percent": "G61",
          "52-Week_value": "H60",
          "52-Week_percent": "H61"
        }
    }
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
