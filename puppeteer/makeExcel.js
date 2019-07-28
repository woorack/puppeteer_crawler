const fs = require('fs');
const excel = require('node-excel-export');
const jsonexport = require('jsonexport');
const json2csv = require('csvjson-json2csv');
const Excel = require('exceljs');
const _ = require('lodash');

// const sample = '[{"enterpriseName":"CÔNG TY TNHH TƯ VẤN THIẾT KẾ KIỂM ĐỊNH XÂY DỰNG N.A.D","numberOfCert":"BXD-00000001","address":"SỐ 22/8 LÊ LỌI - KHU PHỐ 4 - THỊ TRẤN HÓC MÔN - THÀNH PHỐ HỒ CHÍ MINH","provincial":"TP-HCM","lrName":"NGUYỄN NGỌC HOÀI NAM","lrPosition":"GIÁM ĐỐC","regCertNo":"0302461198","regCertDate":"14/11/2001","regCertAgency":"SỞ KẾ HOẠCH VÀ ĐẦU TƯ THÀNH PHỐ HỒ CHÍ MINH","operationList":[{"seqNo":1,"fieldName":"Quản lý dự án đầu tư xây dựng","works":[{"seqNo":1,"work":"Dân dụng","rating":"Hạng 2"},{"seqNo":2,"work":"Công nghiệp","rating":"Hạng 2"},{"seqNo":3,"work":"Giao thông","rating":"Hạng 2"},{"seqNo":4,"work":"Hạ tầng kỹ thuật","rating":"Hạng 2"}]},{"seqNo":2,"fieldName":"Giám sát thi công xây dựng","works":[{"seqNo":1,"work":"Dân dụng","rating":"Hạng 2"},{"seqNo":2,"work":"Công nghiệp","rating":"Hạng 2"},{"seqNo":3,"work":"Giao thông","rating":"Hạng 2"},{"seqNo":4,"work":"Hạ tầng kỹ thuật","rating":"Hạng 2"}]},{"seqNo":3,"fieldName":"Khảo sát xây dựng","works":[{"seqNo":1,"work":"Dân dụng","rating":"Hạng 2"},{"seqNo":2,"work":"Công nghiệp","rating":"Hạng 2"},{"seqNo":3,"work":"Giao thông","rating":"Hạng 2"},{"seqNo":4,"work":"Hạ tầng kỹ thuật","rating":"Hạng 2"}]},{"seqNo":4,"fieldName":"Thiết kế, thẩm tra thiết kế xây dựng","works":[{"seqNo":1,"work":"Dân dụng","rating":"Hạng 2"},{"seqNo":2,"work":"Công nghiệp","rating":"Hạng 2"},{"seqNo":3,"work":"Giao thông","rating":"Hạng 2"},{"seqNo":4,"work":"Hạ tầng kỹ thuật","rating":"Hạng 2"}]},{"seqNo":5,"fieldName":"Quản lý, thẩm tra chi phí đầu tư xây dựng","works":[{"seqNo":1,"work":"Dân dụng","rating":"Hạng 2"},{"seqNo":2,"work":"Công nghiệp","rating":"Hạng 2"},{"seqNo":3,"work":"Giao thông","rating":"Hạng 2"},{"seqNo":4,"work":"Hạ tầng kỹ thuật","rating":"Hạng 2"}]},{"seqNo":6,"fieldName":"Lập, thẩm tra dự án đầu tư xây dựng","works":[{"seqNo":1,"work":"Dân dụng","rating":"Hạng 2"},{"seqNo":2,"work":"Công nghiệp","rating":"Hạng 2"},{"seqNo":3,"work":"Giao thông","rating":"Hạng 2"},{"seqNo":4,"work":"Hạ tầng kỹ thuật","rating":"Hạng 2"}]}]}]';
function flattingJson (json) {
  let result = [];

  json.forEach((enterprise) => {
    if (enterprise !== null) {
      let pickedObj = _.omit(enterprise, 'operationList');
      enterprise.operationList.forEach((operation) => {
        operation.works.forEach((work) => {
          let newRow = Object.assign({}, pickedObj);
          newRow.operationNo = operation.seqNo;
          newRow.operationName = operation.fieldName;
          newRow.workNo = work.seqNo;
          newRow.workName = work.work;
          newRow.workRating = work.rating;

          result.push(newRow);
        });
      });
    }
  });

  return result;
}


let makeExcel = () => {
  let jsonObject = JSON.parse(fs.readFileSync('./results/report_190226.json', 'utf8'));

  try {
    let newJson = flattingJson(jsonObject);
    let output = json2csv(newJson, { flatten: true });
    let workbook = new Excel.Workbook();

    fs.writeFileSync('./results/test.csv', output);
    workbook.csv.readFile('./results/test.csv')
      .then(() => {
        workbook.xlsx.writeFile('./results/test_1.xlsx');
        console.log('Success');
      });

  } catch (err) {
    console.log('ERROR: ', err);
  }

  // const styles = {
  //   headerDark: {
  //     fill: {
  //       fgColor: {
  //         rgb: 'FF000000'
  //       }
  //     },
  //     font: {
  //       color: {
  //         rgb: 'FFFFFFFF'
  //       },
  //       sz: 14,
  //       bold: true,
  //       underline: true
  //     }
  //   }
  // };
  //
  // const spec = {
  //   enterpriseName: {
  //     displayName: 'Name',
  //     headerStyle: styles.headerDark
  //   },
  //   numberOfCert: {
  //     displayName: 'Certificate of Construction activity Capacity No.',
  //     headerStyle: styles.headerDark
  //   },
  //   address: {
  //     displayName: 'Address of head office',
  //     headerStyle: styles.headerDark
  //   },
  //   provincial: {
  //     displayName: 'Province Registered',
  //     headerStyle: styles.headerDark
  //   },
  //   lrName: {
  //     displayName: 'Name of LR',
  //     headerStyle: styles.headerDark
  //   },
  //   lrPosition: {
  //     displayName: 'Position of LR',
  //     headerStyle: styles.headerDark
  //   },
  //   regCertNo: {
  //     displayName: 'Number',
  //     headerStyle: styles.headerDark
  //   },
  //   regCertDate: {
  //     displayName: 'Date',
  //     headerStyle: styles.headerDark
  //   },
  //   regCertAgency: {
  //     displayName: 'Granting Authority',
  //     headerStyle: styles.headerDark
  //   },
  //   operationList: {
  //     displayName: 'TEST',
  //     headerStyle: styles.headerDark
  //   }
  // };
  //
  // const report = excel.buildExport([
  //   {
  //     name: 'Report',
  //     specification: spec,
  //     data: inputFile
  //   }
  // ]);

  // fs.writeFileSync('./results/report.xlsx', report);
};

makeExcel();

