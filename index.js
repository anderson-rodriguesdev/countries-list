// Require library
var xl = require('excel4node');
const axios = require('axios').default;

// Create a new instance of a Workbook class
var wb = new xl.Workbook();
var ws = wb.addWorksheet('Countries List Workbook');

//Create a reusable styles.
var styleHeader = wb.createStyle({
  font: {
    color: '#4F4F4F',
    size: 16,
    bold: true,
  },
  alignment: {
    horizontal: 'center',
  },
});
var styleTitles = wb.createStyle({
  font: {
    color: '#808080',
    size: 12,
    bold: true,
  },
});
var styleNumberFormat = wb.createStyle({
  numberFormat: '#,##0.00',
});

//Create header
function createHeader() {
  //setting tittle inside cell 1-4 with merge 'true'
  ws.cell(1, 1, 1, 4, true).string('Countries List').style(styleHeader);
}

createHeader();

//Create titles
function createTitles() {
  const titles = ['Name', 'Capital', 'Area', 'Currencies'];

  let columnTitles = 1;
  titles.forEach((title) =>
    ws
      .cell(2, columnTitles++)
      .string(title)
      .style(styleTitles),
  );
}

createTitles();

axios({
  method: 'get',
  url: 'https://restcountries.com/v3.1/all',
}).then(function (response) {
  // Listing name
  function listingNames() {
    let columnNames = 3;

    //returning a new array with the names for applying array sort method.
    const nameArray = response.data.map(({ name }) => {
      return name.common;
    });

    nameArray.sort().forEach((item) => {
      ws.cell(columnNames++, 1).string(item);
    });
  }

  listingNames();

  // Listing capital
  function listingCapital() {
    let columnCapital = 3;

    response.data.forEach(({ capital }) => {
      capital
        ? ws.cell(columnCapital++, 2).string(capital)
        : ws.cell(columnCapital++, 2).string('-');
    });
  }

  listingCapital();

  //Listing area
  function listingArea() {
    let columnArea = 3;

    response.data.forEach(({ area }) => {
      area
        ? ws
            .cell(columnArea++, 3)
            .number(area)
            .style(styleNumberFormat)
        : ws.cell(columnArea++, 3).string('-');
    });
  }

  listingArea();

  //Listing Currencies
  function listingCurrencies() {
    let columnCurrencies = 3;

    response.data.forEach(({ currencies }) => {
      currencies
        ? ws
            .cell(columnCurrencies++, 4)
            .string(Object.keys(currencies).toString())
        : ws.cell(columnCurrencies++, 4).string('-');
    });
  }

  listingCurrencies();

  //Create the workbook
  wb.write('countries-list.xlsx');
});
