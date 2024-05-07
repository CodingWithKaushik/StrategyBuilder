import pkg from "xlsx";
import {SMA} from "technicalindicators" ;

const {readFile,utils, writeFile}= pkg;

// Load Excel file and extract data
const workbook = readFile('./124SMA.xlsx');
const sheetName = workbook.SheetNames[0];
const EmtrysheetName = workbook.SheetNames[1];
const TradeLogSheet = workbook.SheetNames[2];
const worksheet = workbook.Sheets[sheetName];
// const data = utils.sheet_to_json(worksheet, { header: ['Date', 'Price'] });
const dateParseOptions = { dateNF: 'yyyy-mm-dd' };

// Extract data from Excel sheet with specified date parsing options
const data = utils.sheet_to_json(worksheet, {
    header: ['Date', 'Price'],
    dateNF: 'yyyy-mm-dd', // Specify date format for parsing 'Date' column
    raw: false // Ensure dates are parsed as JavaScript Date objects
});

// prices
const prices = data.map(entry => entry.Price/1 || 1);

const smaPeriod = 124;
// Define input parameters (close prices for example)
const closePrices = prices;
const smaResult = SMA.calculate({ period: smaPeriod, values: closePrices });

// / Add SMA values to the data
data.forEach((entry, index) => {
    if (index >= smaPeriod - 1) {
        entry['SMA'] = smaResult[index - (smaPeriod - 1)];
        entry['SMARatio'] = (smaResult[index - (smaPeriod - 1)])/entry.Price;
    } else {
        entry['SMA'] = null; // For initial periods where SMA is undefined
    }
});

// Convert updated data back to Excel worksheet
const updatedWorksheet = utils.json_to_sheet(data, { header: ['Date', 'Price', 'SMA','SMARatio'] });

// Update workbook with the updated worksheet
workbook.Sheets[EmtrysheetName] = updatedWorksheet;

// Write the updated workbook back to the same Excel file
writeFile(workbook, './124SMA.xlsx');

 let isEntryActive = false;
 const tradeLog = [];

 data.forEach((entry, index) => {
 if(entry.SMARatio < .92 && !isEntryActive ){

    const entryDetails = {
                        EntryDate: entry.Date,
                        EntryPrice: entry.Price
                    };
                    tradeLog.push(entryDetails);
                    isEntryActive = true;
 }else if (entry.SMARatio > 1 && isEntryActive) {
                const lastTrade = tradeLog[tradeLog.length - 1];
                lastTrade.ExitDate = entry.Date;
                lastTrade.ExitPrice = entry.Price;
                lastTrade.ProfitLoss = lastTrade.ExitPrice - lastTrade.EntryPrice;
                isEntryActive = false;
            }

 })


const tradeLogSheet = utils.json_to_sheet(tradeLog, {
    header: ['EntryDate', 'EntryPrice', 'ExitDate', 'ExitPrice', 'ProfitLoss']
});
workbook.Sheets[TradeLogSheet] = tradeLogSheet;

// Write the updated workbook back to the same Excel file
writeFile(workbook, './124SMA.xlsx');



//30 SMA Entry 
// let isEntryActive = false;
// const tradeLog = [];

// // Loop through data to determine entry and exit points
// data.forEach((entry, index) => {
//     const currentDate = new Date(entry.Date);
//     const currentPrice = entry.Price;
//     const currentSMA = entry.SMA;
//     if (index >= smaPeriod - 1) {
//         const previousSMA = data[index - 1].SMA;
//         const previousEntry = data[index - 1];
//       console.log("CurrentSMA",currentSMA,"PreviousSMA",previousSMA, "currentPrice",currentPrice,"Prev Price", previousEntry.Price);
//         if (currentPrice > previousSMA && !isEntryActive) {
//             // Open new trade (entry)
//             const entryDetails = {
//                 EntryDate: entry.Date,
//                 EntryPrice: currentPrice
//             };
//             tradeLog.push(entryDetails);
//             isEntryActive = true;
//         } else if (currentPrice < currentSMA && isEntryActive) {
//             // Close existing trade (exit)
//             const lastTrade = tradeLog[tradeLog.length - 1];
//             lastTrade.ExitDate = entry.Date;
//             lastTrade.ExitPrice = currentPrice;
//             lastTrade.ProfitLoss = lastTrade.ExitPrice - lastTrade.EntryPrice;
//             isEntryActive = false;
//         }
//     }
// });

// // Write trade log to a new sheet in the workbook
// const tradeLogSheet = utils.json_to_sheet(tradeLog, {
//     header: ['EntryDate', 'EntryPrice', 'ExitDate', 'ExitPrice', 'ProfitLoss']
// });
// workbook.Sheets[TradeLogSheet] = tradeLogSheet;

// // Write the updated workbook back to the same Excel file
// writeFile(workbook, './124SMA.xlsx');

