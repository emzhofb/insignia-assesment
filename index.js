const express = require('express');
const app = express();
const port = 3000;

const axios = require('axios');
const cheerio = require('cheerio');
const Excel = require('exceljs');
const fs = require('fs');

const tokped = {
  item_container: ".pcv3__container",
  item_name: "div .css-1b6t4dn",
  item_price_discount: "div .css-1ksb19c",
  item_price_original: "div .css-1u1z2kp",
  item_rate: "div .css-t70v7i",
  item_sold: "div .css-1duhs3e",
}

app.get('/product/:search', async (req, res, next) => {
  const { search } = req.params;
  const baseUrl = `https://www.tokopedia.com`;
  let url = `${baseUrl}/search?st=product&q=${search.split(' ').join('%20')}`;
  
  try {
    const { data } = await axios.get(url);
    const $ = cheerio.load(data);
    const productGrid = $(tokped.item_container);

    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet('products');
    worksheet.columns = [
      {
        header: 'Name',
        key: 'name',
        width: 100
      },
      {
        header: 'Discount Price',
        key: 'discount_price',
        width: 100
      },
      {
        header: 'Original Price',
        key: 'original_price',
        width: 100
      },
      {
        header: 'Rating',
        key: 'rating',
        width: 100
      },
      {
        header: 'Sold',
        key: 'sold',
        width: 100
      },
    ]
    
    productGrid.each((i, el) => {
      const product = {};
      product.name = $(el).find(tokped.item_name).text();
      product.discount_price = $(el).find(tokped.item_price_discount).text();
      product.original_price = $(el).find(tokped.item_price_original).text() ? $(el).find(tokped.item_price_original).text() : product.discount_price;
      product.rating = $(el).find(tokped.item_rate).text();
      product.sold = $(el).find(tokped.item_sold).text();

      worksheet.addRow(product);
    });

    const path = `${__dirname}/files/${new Date().toISOString()}.xlsx`;
    const excel = await workbook.xlsx.writeBuffer();
    fs.writeFileSync(path, excel);
    res.sendFile(path);
  } catch (error) {
    res.status(400);
  }
});

app.listen(port, () => console.log(`listening on port ${port}`));
