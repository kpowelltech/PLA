// Program that utilizes the xlsx package to work with data in Excel spreadsheets
// Cite: https://www.youtube.com/watch?v=tKz_ryychBY

const xlsx = require("xlsx");
// workbook.SheetNames => shows sheet names

// The file name of the Excel spreadsheet
const savedFile = "PLA_and_AMs.xlsx"

// Loading the spreadsheet into a variable
const workbook = xlsx.readFile(savedFile);

// Loading each spreadsheet within the notebook into its own variable
const PLAs = workbook.Sheets["PLAs"]
const v1 = workbook.Sheets["v1"]
const merchantsAndAMs = workbook.Sheets["Merchants_AMs"]

// Converting spreadsheet data into JSON
const v1Data = xlsx.utils.sheet_to_json(v1)
const plaDATA = xlsx.utils.sheet_to_json(PLAs)

// This logic generates arrays of specific account managers and their merchants
const AMData = xlsx.utils.sheet_to_json(merchantsAndAMs)
const staceyAccounts = []
const aliAccounts = []
const bobbyAccounts = []
const caseyAccounts = []
const christinaAccounts = []
const laurenAccounts = []
const alexAccounts = []

AMData.forEach(account => {
    switch (account.account_manager) {
        case "Stacey Scharton":
            staceyAccounts.push(account);
            break;
        case "Robert McCart":
            bobbyAccounts.push(account);
            break;
        case "Alex Abid":
            alexAccounts.push(account);
            break;
        case "Lauren Levi":
            laurenAccounts.push(account);
            break;
        case "Christina Coy":
            christinaAccounts.push(account);
            break;
        case "Cassandra Penrose":
            caseyAccounts.push(account);
            break;
        case "Ali Rank":
            aliAccounts.push(account);
            break;
    };
});

// The following section handles maping the correct Account Manager to the Merchant with PLA
const completePLA = []

// The function takes in 2 arguments, the Account Manager array of objects and the PLA merchant array of objects
const mapAMtoMerchantHandler = (accountManagerArr, plaArr) => {

    for (let i = 0; i < plaArr.length; i++) {
        for (let j = 0; j < accountManagerArr.length; j++) {
            if (plaArr[i].title === accountManagerArr[j].account_name) {
                plaArr[i].account_manager = accountManagerArr[j].account_manager;
                completePLA.push(plaArr[i])
            }
        }
    }
}

// mapAMtoMerchantHandler(AMData, plaDATA);


// This section is responsibile for creating a new Excel file (workbook) and adding the newly created PLA list into a sheet
const plaWorkbook = xlsx.utils.book_new();
const completedPLAs = xlsx.utils.json_to_sheet(completePLA)
xlsx.utils.book_append_sheet(plaWorkbook, completedPLAs, "PLAs with Merchants")
xlsx.writeFile(plaWorkbook, "PLAs.xlsx");


