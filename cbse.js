const axios = require('axios');
const cheerio = require('cheerio');
const excel = require('exceljs');
const puppeteer = require('puppeteer');
const winston = require('winston');

const logger = winston.createLogger({
  format: winston.format.json(),
  level: 'debug',
  transports: [
    new winston.transports.Console({format: winston.format.simple()}),
  ],
});

const state = process.argv[2];
if (isNaN(parseInt(state))) {
  logger.error('Please run as "node cbse.js <state>" format.');
  process.exit(1);
}

const fields = {
  'Affiliation Number': 'ID',
  'Name of Institution': 'name',
  'Year of Foundation': 'established',
  'Postal Address': 'address',
  'District': 'district',
  'State': 'state',
  'Pin Code': 'pin',
  'Email': 'email',
  'Website': 'website',
  'Name of Principal/ Head of Institution': 'principal',
  'Phone No. with STD Code': 'std',
  'Office': 'office',
  'Residence': 'residence',
};

const trim = (text) => text.trim().replace(/(^,)|(,$)/g, '').trim();

(async () => {
  const browser = await puppeteer.launch();
  const page = await browser.newPage();
  logger.debug('Opening CBSE website...');
  await page.goto('http://saras.cbse.gov.in/cbse_aff/schdir_Report/userview.aspx');
  logger.debug('Clicking on state button...');
  await Promise.all([
    page.click('#optlist_2'),
    page.waitForNavigation({waitUntil: 'networkidle2'}),
  ]);
  logger.debug('Selecting desired state...');
  await Promise.all([
    page.select('#ddlitem', state + ''),
    page.waitForNavigation({waitUntil: 'networkidle2'}),
  ]);
  logger.debug('Initiating search...');
  await Promise.all([
    page.click('#search'),
    page.waitForNavigation({waitUntil: 'networkidle2'}),
  ]);
  let i = 1;
  const rows = [];
  while (true) {
    logger.debug(`Scraping page #${i}...`);
    const tables =
        await page.$$('#T1 table[cellpadding="1"] td:nth-child(2) table');
    logger.debug(`Found ${tables.length} rows...`);
    for (let i = 0; i < tables.length; i++) {
      const first = await tables[i].$$('tbody tr td:first-child');
      const text = await first[0].evaluate((el) => el.textContent);
      const no = /(\d+$)/.exec(text.trim())[1];
      logger.info(`Found affiliation #${no}...`);
      let response;
      try {
        response =
                    await axios.get('http://saras.cbse.gov.in/cbse_aff/schdir_Report/AppViewdir.aspx?affno=' + no);
      } catch (e) {
        logger.error('Failed to fetch affiliation detail from #' + no, e);
      }
      if (response) {
        const $ = cheerio.load(response.data);
        const row = {};
        for (const name in fields) {
          if (fields.hasOwnProperty(name)) {
            const $headings = $('table > tbody > tr > td:first-child');
            $headings.each((i, el) => {
              const $heading = $(el);
              if ($heading.text().indexOf(name) >= 0) {
                const $value = $heading.parent().children().eq(1);
                let value = trim($value.text()).replace('\n', ', ');
                if (['office', 'residence'].indexOf(fields[name]) >= 0) {
                  value = value.split(',').map((x) => trim(x)).join(', ');
                }
                row[fields[name]] = value;
              }
            });
          }
        }
        rows.push(row);
      }
    }
    const disabled = await page.$('#Button1:disabled');
    if (disabled) {
      break;
    }
    logger.debug(`Navigation to next page...`);
    await Promise.all([
      page.click('#Button1'),
      page.waitForNavigation({waitUntil: 'networkidle2'}),
    ]);
    i++;
  }
  await browser.close();
  const workbook = new excel.Workbook();
  const sheet = workbook.addWorksheet('Institutions');
  sheet.addRow([
    'ID',
    'Name',
    'Established in',
    'Street address',
    'District',
    'State',
    'PIN code',
    'Email address',
    'Website URL',
    'Principal',
    'STD code',
    'Office phone',
    'Residence phone',
  ]);
  sheet.addRows(rows.map((row) => [
    row.ID || '',
    row.name || '',
    row.established || '',
    row.address || '',
    row.district || '',
    row.state || '',
    row.pin || '',
    row.email || '',
    row.website || '',
    row.principal || '',
    row.std || '',
    row.office || '',
    row.residence || '',
  ]));
  await workbook.xlsx.writeFile(state + '.xlsx');
})();
