const puppeteer = require('puppeteer');
const moment = require('moment');

async function getPic() {
  const browser = await puppeteer.launch({headless: false});    // headless: without showing browser
  const page = await browser.newPage();
  await page.setViewport({width: 1000, height: 500});
  await page.goto('https://blog.woorack.kr');
  const imagePath = process.cwd() + '/screenshots/blog_' + moment().format('YYYY-MM-YY') + '.png';
  await page.screenshot({path: imagePath});

  await browser.close();
}

getPic();
