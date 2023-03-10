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
                  )[1].includes("Ng???ng Ho???t ?????ng")
                    ? "Ng???ng ho???t ?????ng"
                    : "??ang ho???t ?????ng";
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
      console.log(">> M??? tab m???i...");

      await page.goto(url);
      await page.waitForSelector(
        "body > div > div.col-xs-12.col-sm-9 > div:nth-child(2) > form > div > input"
      );
      await page.focus(
        "body > div > div.col-xs-12.col-sm-9 > div:nth-child(2) > form > div > input"
      );
      await page.type(
        "body > div > div.col-xs-12.col-sm-9 > div:nth-child(2) > form > div > input",
        "k??? to??n"
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
      //   { header: "T??n c??ng ty", key: "companyName", width: 30 },
      //   { header: "T??n qu???c t???", key: "companyNameEn", width: 20 },
      //   { header: "T??n vi???t t???t", key: "companyNameShort", width: 10 },
      //   { header: "M?? s??? thu???", key: "taxID", width: 30 },
      //   { header: "?????a ch???", key: "address", width: 20 },
      //   { header: "Ng?????i ?????i di???n", key: "owners", width: 10 },
      //   { header: "S??? ??i???n tho???i", key: "phone", width: 30 },
      //   { header: "Ng??y ho???t ?????ng", key: "startDay", width: 30 },
      //   { header: "Qu???n l?? b???i", key: "ownerBy", width: 20 },
      //   { header: "Lo???i h??nh DN", key: "legacyType", width: 10 },
      //   { header: "T??nh tr???ng", key: "status", width: 30 },
      // ];
      // arrayA.forEach((row) => {
      //   worksheet.addRow(row);
      // });
      // workbook.xlsx.writeFile("data.xlsx").then(() => {
      //   console.log("???? xu???t t???p Excel th??nh c??ng");
      // });
      resolve();
    } catch (error) {
      console.log("L???i ??? scrape category: " + error);
      reject(error);
    }
  });

module.exports = {
  scrapeCategory,
};
