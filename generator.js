const xlsxFile = require("read-excel-file/node");
const wkhtmltopdf = require("wkhtmltopdf");
const moment = require("moment");
const fs = require("fs");
const _ = require("lodash");

// If you don't have wkhtmltopdf in the PATH, then provide
// the path to the executable (in this case for windows would be):
wkhtmltopdf.command = "C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe";

function run(FILE) {
  let schemaArr = [];
  let schema = {};
  let result = {};

  const ROOT = "C:/Users/Akhil/Desktop/WebDev";
  let staff = {
    teacher: "Dayanand Sharma",
    coordinator: "Dr. Alok Saxena",
    principal: "Girdhar Kumari",
  };
  const ID = "Sch. No.";
  const NAME = "Name of Student";
  const CLASS = "Class";
  const rollNo = "Roll No";
  const demographic = [
    "Roll No",
    "Sch. No.",
    "Name of Student",
    "Class",
    "Phone No.",
    "Email id",
  ];
  let subjects = [];

  const getData = (row, col) => {
    return row[schema[col]];
  };

  const generatePDF = (row) => {
    const fileName = `${getData(row, rollNo)}_${getData(row, NAME).trim().replace(' ', '_')}.pdf`;
    const grade = getData(row, CLASS);
    const student = demographic.map((e) => ({
      label: e,
      value: getData(row, e),
    }));
    const marks = subjects.map((e) => ({
      label: e.name,
      value: row[e.index],
      mm: e.mm,
    }));
    const htmlContent = `
      <!DOCTYPE html>
      <html>
          <head>
              <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.0/css/bootstrap.min.css" integrity="sha384-9aIt2nRpC12Uk9gS9baDl411NQApFmC26EwAOH8WgZl5MYYxFfc+NcPb1dKGj7Sk" crossorigin="anonymous">
              <style>
                  .container {
      
                  }
                  .details {
                      padding-top: 32px;
                  }
                  .details .item {
      
                  }
                  .marks {
                      padding-top: 32px;
                  }
                  .teachers {
                      padding-top: 64px;
                  }
                  .footer {
                      padding-top: 64px;
                  }
                  table, tr, td, th, tbody, thead, tfoot {
                      page-break-inside: avoid !important;
                  }
                  .center-text {
                      text-align: center;
                  }
                  .flex-row {
                      display: flex;
                      justify-content: space-between;
                  }
              </style>
          </head>
          <body>
              <div class="container d-flex flex-column justify-content-around">
                  <section class="heading">
                      <div class="center-text">
                          <img src="file:///C:/Users/Akhil/Desktop/WebDev/assets/images/sanskar.png" alt="Sanskar" height="150px">
                      </div>
                      <h3 class="pt-4 center-text text-uppercase">Unit Test 1 - May 2020</h3>
                  </section>
                  <section class="details">
                      <h4>Student Details:</h4>
                      <table class="table table-bordered table-striped table-sm">
                          <thead>
                          </thead>
                          <tbody>
                              ${student
                                .map((o) => {
                                  return `
                                          <tr>
                                              <th scope="row">${
                                                o.label
                                              }</th>    
                                              <td>${
                                                o.value || "Not Available"
                                              }</td>
                                          </tr>`;
                                })
                                .join("")}
                          </tbody>
                      </table>
                  </section>
                  <section class="marks">
                      <h4>Student Marks:</h4>
                      <table class="table table-bordered table-striped table-sm">
                          <thead>
                              <tr>
                                  <th>Subject</th>
                                  <th>Maximum Marks</th>
                                  <th>Marks</th>
                              </tr>
                          </thead>
                          <tbody>
                              ${marks
                                .map((o) => {
                                  return `
                                          <tr>
                                              <th scope="row">${o.label}</th>
                                              <td>${o.mm}</td>
                                              <td>${
                                                o.value || "Not Available"
                                              }</td>
                                          </tr>`;
                                })
                                .join("")}
                          </tbody>
                      </table>
                  </section>
                  <section class="teachers">
                      <table class="table table-borderless table-sm">
                          <tbody>
                              <tr>
                                  <td>${_.startCase(
                                    _.toLower(staff.teacher)
                                  )}</td>
                                  <td>${staff.coordinator}</td>
                                  <td>${staff.principal}</td>
                              </tr>
                              <tr>
                                  <td class="font-weight-bold">Class Teacher</td>
                                  <td class="font-weight-bold">Coordinator</td>
                                  <td class="font-weight-bold">Principal</td>
                              </tr>
                          </tbody>
                      </table>
                  </section>
                  <section class="footer center-text">
                      <div class="">
                          <img src="file:///C:/Users/Akhil/Desktop/WebDev/assets/images/cambridge.png" alt="Cambridge" height="100px">
                      </div>
                      <p class="text-monospace">This is auto-generated result, please contact the coordinator for any discrepancies.</p>
                      <div class="p-2 text-monospace date">
                          <span class="font-weight-bold">Date: </span>
                          <span>${moment().format("DD/MM/YYYY")}</span>
                      </div>
                  </section>
              </div>
          </body>
      </html>
      `;
    const dir = `${ROOT}/results/${grade}`;
    const URI = `${dir}/${fileName}`;
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir);
    }
    wkhtmltopdf(htmlContent, {
      output: URI,
      pageSize: "A4",
    });
    return URI;
  };

  return xlsxFile(`${ROOT}/data/${FILE}`).then((rows) => {
    let isHeaderSet = false;
    let isMMset = false;
    rows.forEach((row, i) => {
      const rowString = row.join(" ").trim().toLowerCase();
      const identifier = "teacher-";
      const match = rowString.indexOf(identifier);
      if (rowString === "") {
        return;
      }
      if (match > -1) {
        staff.teacher = rowString.substr(match + identifier.length);
        return;
      }
      if (!isHeaderSet) {
        console.log(row);
        row.forEach((col, j) => {
          schema[col] = j;
          schemaArr.push(col);
        });
        isHeaderSet = true;
        return;
      }
      if (!isMMset) {
        let subj = [];
        row.forEach((e, i) => {
          const numberPattern = /M.M.\s*(\d+)/;
          const matchMM = (e || "").match(numberPattern);
          if (matchMM && matchMM[1]) {
            subj.push({
              name: schemaArr[i],
              index: i,
              mm: matchMM[1],
            });
          }
        });
        subjects = subj;
        isMMset = true;
        return;
      }
      const URI = generatePDF(row);
      const id = getData(row, ID);
      result[id] = URI;
      return;
    });
    console.log(result);
    return result;
  });
}

module.exports = {
  run
}
