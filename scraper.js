const ExcelJS = require("exceljs");
const scrapeCategory = async (browser, url) =>
  new Promise(async (resolve, reject) => {
    try {
      let taxID;
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

      await page.waitForNavigation();
      const currentUrl = page.url();
      await page.waitForTimeout(1000);
      const lastPageUrl = await page.$eval(
        "body > div > div.col-xs-12.col-sm-9 > ul > li:nth-child(6) > a",
        (anchor) => anchor.getAttribute("href")
      );
      console.log(lastPageUrl);
      const params = new URL(lastPageUrl);
      const total = Number(params.searchParams.get("page"));
      console.log(typeof total); // Output: "22"

      for (let k = 1; k <= total; k++) {
        if (k === 1) {
        } else {
          await page.goto(currentUrl.concat(`/?page=${k}`));

          await page.waitForNavigation();
        }

        let taxListing = await page.$$(
          "body > div.container > div.col-xs-12.col-sm-9 > div:nth-child(3) > div"
        );

        for (let j = 2; j <= taxListing.length; j++) {
          if (j === 6 || j === 28) {
            continue;
          } else {
            const Id = await page.$$eval(
              `body > div.container > div.col-xs-12.col-sm-9 > div:nth-child(3) > div:nth-child(${j}) > p`,
              (trs) => {
                let result = "";
                Array.from(trs, (tr) => {
                  const th = tr.querySelector("a");
                  result = th.innerText?.trim();
                });
                return result;
              }
            );
            taxID = Id;
          }
          await page.goto("https://masothue.com/");
          await page.waitForSelector("#search");
          await page.focus("#search");
          await page.type("#search", taxID);
          await page.click(
            "#page > nav > div > form > div > div.input-group-btn > button"
          );

          await page.waitForNavigation();

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
          arrayA.push(dataReusltInfoCompany);
          console.log(dataReusltInfoCompany);
          await page.goto(currentUrl.concat(`&page=${k}`));
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
