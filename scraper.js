const ExcelJS = require("exceljs");

const scrapeCategory = async (browser, url, taxID) =>
  new Promise(async (resolve, reject) => {
    try {
      if (taxID) {
        console.log(taxID);

        let page = await browser.newPage();
        await page.goto("https://masothue.com");

        await page.waitForSelector(".tax-search input.search-field");

        await page.type(".tax-search input.search-field", taxID);

        await page.keyboard.press("Enter");

        await page.waitForSelector(".table-taxinfo");

        const companyName = await page.$$eval(
          ".table-taxinfo thead tr",
          (trs) => {
            let result = "";
            Array.from(trs, (tr) => {
              const th = tr.querySelector("th");
              result = th.innerText?.trim();
            });
            return result;
          }
        );

        const dataReusltInfoCompany = await page.$$eval(
          ".table-taxinfo tbody tr",
          (trs) => {
            const reusltInfoCompany = {};
            Array.from(trs, (tr) => {
              const columns = [
                "companyNameEn",
                "companyNameShort",
                "taxID",
                "address",
                "owners",
                "phone",
                "startDay",
                "ownerBy",
                "legacyType",
                "status",
              ];

              const capitalizeFirstLetterOfEachWord = (str) => {
                let splitStr = str?.toLowerCase()?.split(" ");
                for (let i = 0; i < splitStr?.length; i++) {
                  splitStr[i] =
                    splitStr[i].charAt(0).toUpperCase() +
                    splitStr[i].substring(1);
                }
                return splitStr?.join(" ");
              };

              const tds = tr.querySelectorAll("td");

              const eleI = tds[0].querySelector("i");
              const text = tds[1]?.innerText?.trim();
              const classNameOfEleI = eleI?.className;

              switch (classNameOfEleI) {
                case "fa fa-globe":
                  reusltInfoCompany[columns[0]] =
                    capitalizeFirstLetterOfEachWord(text);
                case "fa fa-reorder":
                  reusltInfoCompany[columns[1]] = text;
                case "fa fa-hashtag":
                  reusltInfoCompany[columns[2]] = text;
                case "fa fa-map-marker":
                  reusltInfoCompany[columns[3]] =
                    capitalizeFirstLetterOfEachWord(text);
                case "fa fa-user":
                  reusltInfoCompany[columns[4]] =
                    capitalizeFirstLetterOfEachWord(text);
                case "fa fa-phone":
                  reusltInfoCompany[columns[5]] = text;
                case "fa fa-calendar":
                  let contentArrayCalendar = Array.from(tds, (td) =>
                    td?.innerText?.trim()
                  );
                  const contentSplit = contentArrayCalendar[1]?.split("-");
                  reusltInfoCompany[
                    columns[6]
                  ] = `${contentSplit[2]}/${contentSplit[1]}/${contentSplit[0]}`;
                case "fa fa-users":
                  reusltInfoCompany[columns[7]] =
                    capitalizeFirstLetterOfEachWord(text);
                case "fa fa-building":
                  reusltInfoCompany[columns[8]] = text;
                case "fa fa-info":
                  reusltInfoCompany[columns[9]] = Array.from(tds, (td) =>
                    td?.innerText?.trim()
                  )[1].includes("Ngừng Hoạt Động")
                    ? "Ngừng hoạt động"
                    : "Đang hoạt động";
                default:
              }
            });

            return reusltInfoCompany;
          }
        );
        dataReusltInfoCompany.companyName = companyName;
        // console.log(dataReusltInfoCompany);
        await page.close();

        return resolve(dataReusltInfoCompany);
      }
      let B;
      const arrayA = [];
      let page = await browser.newPage();
      console.log(">> Mở tab mới...");

      await page.goto(url);
      await page.waitForSelector(
        "body > div > div.col-xs-12.col-sm-9 > div:nth-child(2) > form > div > input"
      );
      await page.focus(
        "body > div > div.col-xs-12.col-sm-9 > div:nth-child(2) > form > div > input"
      );
      await page.type(
        "body > div > div.col-xs-12.col-sm-9 > div:nth-child(2) > form > div > input",
        "kế toán"
      );
      await page.click(
        "body > div > div.col-xs-12.col-sm-9 > div:nth-child(2) > form > div > div > button"
      );

      await page.waitForSelector(".search-results");

      const currentUrl = page.url();

      const lastPageUrl = await page.$eval(
        "body > div > div.col-xs-12.col-sm-9 > ul > li:nth-child(6) > a",
        (anchor) => anchor.getAttribute("href")
      );

      const params = new URL(lastPageUrl);

      const total = Number(params.searchParams.get("page"));

      for (let k = 1; k <= total; k++) {
        await page.goto(currentUrl.concat(`?page=${k}`));

        await page.waitForSelector("div.search-results");

        const taxListing = await page.$$eval("div.search-results", (results) =>
          results.map((result) => {
            return result.querySelector("p a").innerText?.trim();
          })
        );

        for (let j = 0; j < taxListing.length; j++) {
          B = await scrapeCategory(
            browser,
            "https://masothue.com",
            taxListing[j]
          );
          arrayA.push(B);
          await page.waitForTimeout(1500);
        }
      }
      // const workbook = new ExcelJS.Workbook();
      // const worksheet = workbook.addWorksheet("Sheet1");
      // worksheet.columns = [
      //   { header: "Tên công ty", key: "companyName", width: 30 },
      //   { header: "Tên quốc tế", key: "companyNameEn", width: 20 },
      //   { header: "Tên viết tắt", key: "companyNameShort", width: 10 },
      //   { header: "Mã số thuế", key: "taxID", width: 30 },
      //   { header: "Địa chỉ", key: "address", width: 20 },
      //   { header: "Người đại diện", key: "owners", width: 10 },
      //   { header: "Số Điện thoại", key: "phone", width: 30 },
      //   { header: "Ngày hoạt động", key: "startDay", width: 30 },
      //   { header: "Quản lý bởi", key: "ownerBy", width: 20 },
      //   { header: "Loại hình DN", key: "legacyType", width: 10 },
      //   { header: "Tình trạng", key: "status", width: 30 },
      // ];
      // arrayA.forEach((row) => {
      //   worksheet.addRow(row);
      // });
      // workbook.xlsx.writeFile("data.xlsx").then(() => {
      //   console.log("Đã xuất tệp Excel thành công");
      // });
      resolve();
    } catch (error) {
      console.log("Lỗi ở scrape category: " + error);
      reject(error);
    }
  });

module.exports = {
  scrapeCategory,
};
