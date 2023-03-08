const tes = async () => {
  const category = await sequelize.query("SELECT * FROM product_category", {
    type: QueryTypes.SELECT,
  });

  const detail = await sequelize.query("SELECT * FROM product ORDER BY id ASC", {
    type: QueryTypes.SELECT,
  });

  let data = [];

  for (var i = 0; i < category.length; i++) {
    const item = {
      category_id: category[i].id,
      category: category[i].name,
      detail: detail.filter((itm) => {
        return itm.category_id === category[i].id;
      }),
    };
    data.push(item);
  }

  const workbook = new Excel.Workbook();
  const name = `new-${uuid.v4()}.xlsx`;
  workbook.xlsx.readFile("template.xlsx").then(function () {
    var worksheet = workbook.getWorksheet(1);

    for (var x = 0; x < data.length; x++) {
      var rowFirst = [];
      rowFirst[1] = x + 1;
      rowFirst[3] = data[x].category;
      const roow = 7;

      for (var i = 0; i < data[x].detail.length; i++) {
        var rowValues = [];
        rowValues[1] = `${x + 1}.${i + 1}`;
        rowValues[3] = data[x].detail[i].name;
        rowValues[4] = "F";
        rowValues[8] = data[x].detail[i].unit_price;
        worksheet.insertRow(7, rowValues, "o");
      }
      worksheet.insertRow(roow, rowFirst, "o");
    }

    return workbook.xlsx.writeFile(name);
  });
};
tes();
