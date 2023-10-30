const ExcelJS = require('exceljs');
const fs = require('fs');
const cheerio = require('cheerio');

const Config = {
  HTML_FILE_PATH: '',
  XLSX_FILE_PATH: '',
  WORKBOOK_NAME: '',
  COLUMN: '',
  ROW: number,
  ROW_INTERVAL: number,
  CSS_SELECTOR: '',
  CLEAR_CHILDREN: false,
};

const changeText = (config) => {
  const workbook = new ExcelJS.Workbook();

  workbook.xlsx.readFile(config.XLSX_FILE_PATH)
    .then(() => {
      const worksheet = workbook.getWorksheet(config.WORKBOOK_NAME);

      fs.readFile(config.HTML_FILE_PATH, 'utf-8', (err, data) => {
        if (err) throw err;
        const htmlFileData = cheerio.load(data);

        const htmlElements = htmlFileData(config.CSS_SELECTOR);

        htmlElements.each((index, element) => {
          const htmlElement = htmlFileData(element);
          const cell = worksheet.getCell(config.COLUMN + config.ROW);
          const replacedText = cell.value;

          if (!config.CLEAR_CHILDREN) {
            const children = htmlElement.children().toArray();
            const tempContainer = htmlFileData('<div></div>');
            tempContainer.append(children);
            htmlElement.empty().text(replacedText);
            htmlElement.append(tempContainer.html());
          } else {
            htmlElement.empty().text(replacedText);
          }

          config.ROW += config.ROW_INTERVAL;
        });

        const updatedHtml = htmlFileData.html();

        fs.writeFile(config.HTML_FILE_PATH, updatedHtml, 'utf-8', (err) => {
          if (err) throw err;

          console.log('Готово!');
        });
      });
    });
};

changeText(Config);
