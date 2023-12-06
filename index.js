const axios = require('axios');
const cheerio = require('cheerio');
const excelJS = require('exceljs');
const {Builder, By, Capabilities, until} = require('selenium-webdriver');
const chrome = require('selenium-webdriver/chrome');
const fs = require('fs');


const url = 'https://www.barchart.com/futures/quotes/SBH24/performance';

const sheet_map = {
    0: 'E12',
    1: 'F12',
    2: 'G12'
}

async function getLastPrice() {
    let chromeCapabilities = Capabilities.chrome();
    let chromeOptions = new chrome.Options();

    // Set preferences to disable redirects
    chromeOptions.setUserPreferences({ 'profile.default_content_setting_values.automatic_downloads': 2 })
    chromeOptions.setPageLoadStrategy('eager')
    let driver = await new Builder().forBrowser('chrome').setChromeOptions(chromeOptions)
    .withCapabilities(chromeCapabilities).build();
    try {
        // Navigate to the URL
        await driver.get('https://www.barchart.com/futures/quotes/SBH24/performance')
       
        // Wait for the element containing the last price to be loaded
        let lastPriceElement = await driver.wait(until.elementLocated(By.className('pricechangerow')), 10000);

        let lastPrice = await lastPriceElement.getText();
        console.log('Last Price:', lastPrice);
    } finally {
        // Close the browser
        await driver.quit();
    }
}

getLastPrice();



// axios.get(url)
//   .then(async response => {
//     const html = response.data;
//     const $ = cheerio.load(html);

//     // const workbook = new excelJS.Workbook();
//     // await workbook.xlsx.readFile('./DAILY-REPORT.xlsx');
//     // const worksheet = workbook.getWorksheet('Thế giới');
//      //     if (sheet_map?.[i]) {
//     //         worksheet.getCell(sheet_map[i]).value = period_value_arr[i];
//     //     }
//     const priceChange = $('.pricechangerow');
//     const item = await scrapeData(url);
//     console.log('item', item);


//     // const performanceTable = $('.block-content');
//     // performanceTable.each(function() {
//     //   // Extract and process the required data
//     //   let period = $(this).find('td.cell-period-change-title'); 
//     //   let period_value = $(this).find('td.cell-period-change');
//     //   const period_arr = period.map(function (i, el) {
//     //     return $(this).text().trim();
//     //   })
//     //   .toArray();
//     //   const period_value_arr = period_value.map(function (i, el) {
//     //     let spanValue = $(this).children().find('span').first();
//     //     let spanPercent = $(this).children().find('span').last();
//     //     return { value: spanValue.text(), percent: spanPercent.text().replace(/[()]/g, '')  };
//     //   })
//     //   .toArray();
//     //   const totalMap = period_arr.map((e, i) => {
//     //     return {
//     //         preiod: e,
//     //         value: period_value_arr[i].value,
//     //         percent: period_value_arr[i].percent,
//     //     }
//     //   });
//     //   console.log('totalMap', totalMap);
//     // });

//     // await workbook.xlsx.writeFile('./DAILY-REPORT.xlsx');
//   })
//   .catch(console.error);


