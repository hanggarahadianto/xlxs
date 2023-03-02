const express = require("express");
const app = express();

const util = require("util");
const fs = require("fs");
const Excel = require('exceljs');
const path = require("path");
const uuid = require("uuid");

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

const tes = () => {
  const workbook = new Excel.Workbook();
  const name = `new-${uuid.v4()}.xlsx`;
    workbook.xlsx.readFile('old.xlsx')
    .then(function() {
        var worksheet = workbook.getWorksheet(1);
        var row = worksheet.getRow(1);
        row.getCell(4).value = "Tes Judul Excel"; 
        row.getCell(8).value = "no/22/001"; 
        row.commit();

        var row2 = worksheet.getRow(2);
        row2.getCell(8).value = "1 maret 2023"; 
        row2.commit();


        var row7 = worksheet.getRow(7);
        row7.getCell(1).value = "1.1"; 
        row7.getCell(3).value = "Bahan dan alat tertata rapi"; 
        row7.getCell(4).value = "F"; 
        row7.getCell(8).value = "10"; 
        
        row7.commit();
        return workbook.xlsx.writeFile(name);
    })
  };
  tes();

app.get("/", (req, res) => {
  res.render("index", { title: "convert" });
});

app.listen(3000, () => {
  console.log("Server started on port 3000");
});
