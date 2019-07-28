const puppeteer = require('puppeteer');
const moment = require('moment');
const json2xls = require('json2xls');
const fs = require('fs');
const excel = require('node-excel-export');

let scrape = async () => {
  const browser = await puppeteer.launch({headless: true});
  const page = await browser.newPage();

  try {
    const result = [];

    // Scrape
    // There's not enterprise data Before no.1914.
    for (let index = 1914; index < 12791; index++) {
      let targetUrl = `http://hdxd.xaydung.gov.vn/cqlhdxd/xem-thong-tin-nha-thau/nha-thau-${index}.html`;
      console.log('>>', targetUrl);
      await page.goto(targetUrl, {timeout: 0});
      await page.addScriptTag({ url: 'https://code.jquery.com/jquery-3.2.1.min.js' });

      let value = await page.evaluate(() => {
        let enterpriseName = document.querySelectorAll('.col-sm-6.xuongdong')[0].innerText;
        if (enterpriseName) {
          let values = {
            enterpriseName,
            numberOfCert: document.querySelector('.col-sm-10.xuongdong').innerText,
            address: document.querySelectorAll('.col-sm-6.xuongdong')[1].innerText,
            provincial: document.querySelectorAll('.col-sm-2.xuongdong')[1].innerText,
            lrName: document.querySelectorAll('.col-sm-2.xuongdong')[2].innerText,
            lrPosition: document.querySelectorAll('.col-sm-2.xuongdong')[3].innerText,
            regCertNo: document.querySelectorAll('.col-sm-2.xuongdong')[6].innerText,
            regCertDate: document.querySelectorAll('.col-sm-2.xuongdong')[7].innerText,
            regCertAgency: document.querySelectorAll('.col-sm-3.xuongdong')[0].innerText
          };

          let operationDoms = $('.wrap-frm.wrap-frm_');
          let operationList = [];
          for(let index = 0; index < operationDoms.length; index++) {
            let tempWork = {
              seqNo: index + 1,
              fieldName: operationDoms[index].children[1].innerText,
              works: []
            };

            let workList = operationDoms[index].children[2].children;
            for(let jndex = 0; jndex < workList.length; jndex++) {
              let works = {
                seqNo: jndex + 1,
                work: workList[jndex].children[0].innerText,
                rating: workList[jndex].children[1].innerText
              };
              tempWork.works.push(works);
            }
            operationList.push(tempWork);
          }
          values.operationList = operationList;

          return values;
        }
      });

      result.push(value);
    }
    browser.close();

    return result;
  } catch (err) {
    console.log(err);
    throw err;
  }
};

scrape().then((value) => {
  fs.writeFileSync('./results/report_190226.json', JSON.stringify(value));
});