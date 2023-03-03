const express = require("express");
const app = express();

const util = require("util");
const fs = require("fs");
const Excel = require("exceljs");
const path = require("path");
const uuid = require("uuid");

app.set("view engine", "ejs");

const XLSX = require("xlsx");

var isJson = require("is-json");

var json2xls = require("json2xls");

var bodyParser = require("body-parser");

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: false }));

const Sequelize = require("sequelize");

const sequelize = new Sequelize("Questioner", null, null, {
  username: "root",
  password: "12345678",
  database: "Questioner",
  host: "localhost",
  dialect: "mysql",
  logging: false,
  dialectOptions: {
    options: {
      requestTimeout: 1000000,
      connectTimeout: 1000000,
    },
  },
});
const { QueryTypes } = require("sequelize");

const tes = async () => {
  const submission_details = await sequelize.query("SELECT * FROM submission WHERE doc_no = 'FM-HSE-13'", {
    type: QueryTypes.SELECT,
  });

  const answer_details = await sequelize.query("SELECT * FROM answer WHERE id = 56", {
    type: QueryTypes.SELECT,
  });

  const question_details = await sequelize.query("SELECT * FROM question WHERE id = 16", {
    type: QueryTypes.SELECT,
  });

  const form_details = await sequelize.query("SELECT * FROM form WHERE id = 6", {
    type: QueryTypes.SELECT,
  });

  const workbook = new Excel.Workbook();

  const name = `new-${uuid.v4()}.xlsx`;
  workbook.xlsx.readFile("old.xlsx").then(function () {
    var worksheet = workbook.getWorksheet(1);

    var row = worksheet.getRow(1);
    row.getCell(4).value = form_details[0].nama_form;
    row.getCell(8).value = submission_details[0].doc_no;
    row.commit();

    var row2 = worksheet.getRow(2);
    row2.getCell(8).value = submission_details[0].tanggal_dibuat;
    row2.commit();

    var row3 = worksheet.getRow(3);
    row3.getCell(8).value = submission_details[0].tanggal_revisi;
    row3.commit();

    var row4 = worksheet.getRow(4);
    row4.getCell(8).value = submission_details[0].nama_cabang;
    row4.commit();

    var row7 = worksheet.getRow(7);
    row7.getCell(1).value = "1.1";
    row7.getCell(2).value = question_details[0].indikator;
    row7.getCell(4).value = question_details[0].metode_verifikasi;
    row7.getCell(5).value = answer_details[0].metode_verifikasi;
    row7.getCell(7).value = answer_details[0].nilai;
    row7.getCell(8).value = question_details[0].bobot;
    row7.getCell(9).value = answer_details[0].keterangan;
    row7.commit();

    return workbook.xlsx.writeFile(name);
  });
};
tes();
