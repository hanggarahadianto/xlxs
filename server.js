const express = require("express");
const app = express();

const Excel = require("exceljs");

const uuid = require("uuid");

app.set("view engine", "ejs");

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

const tes = async (req) => {
  let submission_details = await sequelize.query(
    `SELECT 
    submission.id,
    submission.doc_no,
    directus_users.first_name,
    directus_users.last_name,
    form.nama_form,
    form_category.nama_kategori,
    form_category.id AS id_kategori,

    submission.tanggal_dibuat,
    submission.tanggal_revisi,
    submission.nama_cabang,
    
    question.indikator,
    question.metode_verifikasi,
    question.bobot,
    
    answer.hasil,
    answer.nilai,
    answer.keterangan

  FROM submission

    INNER JOIN form ON submission.nama_form=form.id 
    INNER JOIN directus_users ON submission.user = directus_users.id
    INNER JOIN form_category ON form_category.category_id=form.id
    INNER JOIN question ON question.question_id=form_category.id
    LEFT JOIN answer ON answer.question_id=question.id
   
   WHERE doc_no = "${req.query.doc_no}"`,

    {
      type: QueryTypes.SELECT,
    }
  );

  console.log(submission_details);

  let data = [];

  const groupByCategory = submission_details.reduce((group, item) => {
    const { nama_kategori } = item;
    group[nama_kategori] = group[nama_kategori] ?? [];
    group[nama_kategori].push(item);
    return group;
  }, {});

  for (let key in groupByCategory) {
    data.push({
      id_kategori: groupByCategory[key][0].id_kategori,
      nama_kategori: key,
      detail: groupByCategory[key],
    });
  }

  const workbook = new Excel.Workbook();

  const name = `new-${uuid.v4()}.xlsx`;
  workbook.xlsx.readFile("old.xlsx").then(function () {
    var worksheet = workbook.getWorksheet(1);

    const style = {
      name: "Comic Sans MS",
      family: 4,
      size: 10,
      underline: true,
      bold: true,
    };

    var row = worksheet.getRow(1);
    row.getCell(2).value = submission_details[0].nama_form;
    row.getCell(2).font = style;
    row.getCell(4).value = submission_details[0].doc_no;
    row.getCell(4).font = style;

    row.commit();

    var row2 = worksheet.getRow(2);
    row2.getCell(4).value = submission_details[0].tanggal_dibuat;
    row2.getCell(4).font = style;
    row2.commit();

    var row3 = worksheet.getRow(3);
    row3.getCell(4).value = submission_details[0].tanggal_revisi;
    row2.getCell(4).font = style;
    row3.commit();

    var row4 = worksheet.getRow(4);
    row4.getCell(4).value = submission_details[0].nama_cabang;
    row4.commit();

    for (var x = data.length - 1; x >= 0; x--) {
      const rowFirst = [];

      rowFirst[1] = x + 1;
      rowFirst[2] = data[x].nama_kategori;

      const row = 7;

      for (var i = data[x].detail.length - 1; i >= 0; i--) {
        var rowValues = [];
        rowValues[1] = `${x + 1}.${i + 1}`;
        rowValues[2] = data[x].detail[i].indikator;

        rowValues[3] = data[x].detail[i].metode_verifikasi;
        rowValues[4] = data[x].detail[i].hasil;
        rowValues[5] = data[x].detail[i].nilai;
        rowValues[6] = data[x].detail[i].bobot;
        rowValues[7] = data[x].detail[i].keterangan;

        worksheet.insertRow(7, rowValues, "o");
      }
      worksheet.insertRow(row, rowFirst, "o");
    }

    return workbook.xlsx.writeFile(name);
  });
};

app.get("/create-excel", (req, res) => {
  tes(req);
  console.log(tes);
});

app.listen(3000, () => {
  console.log("Server started on port 3000");
});
