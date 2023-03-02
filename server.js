const express = require("express");
const app = express();

const util = require("util");
const fs = require("fs");
const conversionFactory = require("html-to-xlsx");
const puppeteer = require("puppeteer");
const chromeEval = require("chrome-page-eval")({ puppeteer });
const writeFileAsync = util.promisify(fs.writeFile);
const path = require("path");

app.set("view engine", "ejs");

const XLSX = require("xlsx");

var isJson = require("is-json");

var json2xls = require("json2xls");

var bodyParser = require("body-parser");

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: false }));

const questionnaire = [
  {
    // date_created: "2023-03-01T07:01:44.000Z",
    // date_updated: "2023-03-01T07:01:44.000Z",
    // id: 98,
    // sort: null,
    doc_no: "FM-HSE-11",
    nama_form: 6,
    nama_cabang: null,
    tanggal_dibuat: "2023-03-01T07:01:44",
    tanggal_revisi: null,
    user: "b452df24-a638-44a0-8cbf-a71647f28234",
    user_created: "16c150fe-3d8b-42d3-81d7-9fc2cc227761",
    user_updated: "16c150fe-3d8b-42d3-81d7-9fc2cc227761",
    answer_details: [
      {
        question_id: 5,
        nilai: 10,
        keterangan: "bagus",
      },
    ],
  },
];

// const convertJsonToExcel = () => {
//   const workSheet = XLSX.utils.json_to_sheet(questionnaire);
//   const workBook = XLSX.utils.book_new();

//   XLSX.utils.book_append_sheet(workBook, workSheet, "questionnaire");
//   // Generate buffer
//   XLSX.write(workBook, { bookType: "xlsx", type: "buffer" });

//   // Binary string
//   XLSX.write(workBook, { bookType: "xlsx", type: "binary" });

//   XLSX.writeFile(workBook, "questionnaireData.xlsx");
// };
// convertJsonToExcel();

const conversion = conversionFactory({
  extract: async ({ html, ...restOptions }) => {
    const tmpHtmlPath = path.join("", "input.html");

    await writeFileAsync(tmpHtmlPath, html);

    const result = await chromeEval({
      ...restOptions,
      html: tmpHtmlPath,
      scriptFn: conversionFactory.getScriptFn(),
    });

    const tables = Array.isArray(result) ? result : [result];

    return tables.map((table) => ({
      name: table.name,
      getRows: async (rowCb) => {
        table.rows.forEach((row) => {
          rowCb(row);
        });
      },
      rowsCount: table.rows.length,
    }));
  },
});

async function run() {
  const stream = await conversion(`<table><tr><td>cell value</td></tr></table>`);

  stream.pipe(fs.createWriteStream("output.xlsx"));
}

run();

app.get("/", (req, res) => {
  res.render("index", { title: "convert" });
});

app.listen(3000, () => {
  console.log("Server started on port 3000");
});