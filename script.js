function handleFileChange(event) {
  const file = event.target.files[0];
  if (!/\.(xlsx|xls)$/.test(file.name)) {
    alert("Please select an Excel file (.xlsx or .xls)");
    return;
  }
  readExcelFile(file);
}

function readExcelFile(file) {
  const reader = new FileReader();
  reader.onload = function (event) {
    const data = event.target.result;
    const workbook = XLSX.read(data, { type: "binary" });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    let jsonData = XLSX.utils.sheet_to_json(worksheet);
    jsonData.forEach((item) => {
      if (item.hasOwnProperty("article_date")) {
        const rawDate = item.article_date;
        const baseDate = new Date(1900, 0, 0);
        const formattedDate = new Date(baseDate.getTime() + rawDate * 86400000);
        item.article_date = formattedDate.toISOString().slice(0, 10);
      }
    });
    jsonData = groupCompetitors(jsonData);
    processExcelData(jsonData);
  };
  reader.readAsArrayBuffer(file);
}

function groupCompetitors(jsonData) {
  const result = [];
  let competitorInfo = null;

  for (const item of jsonData) {
    if (item.category === "Competitor") {
      if (!competitorInfo) {
        competitorInfo = { category: item.category, articles: [] };
      }
      competitorInfo.articles.push({
        organisation_name: item.organisation_name,
        article_title: item.article_title,
        article_description: item.article_description,
        article_link: item.article_link,
        article_date: item.article_date,
      });
    } else {
      result.push(item);
    }
  }

  if (competitorInfo) {
    result.push(competitorInfo);
  }
  return result;
}

function processExcelData(jsonData) {
  let emailTemplate = generateEmailTemplate(jsonData);
  displayTemplate(emailTemplate);
  downloadTemplate(emailTemplate);
}

function formatDate(dateString) {
  const [year, month, day] = dateString.split("-");
  return `${day.padStart(2, "0")}-${month.padStart(2, "0")}-${year}`;
}

function generateEmailTemplate(jsonData) {
  // console.log(jsonData);
  let emailTemplate = `<!DOCTYPE html>
      <html lang="en">
      <head>
        <meta charset="UTF-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
        <link href="https://fonts.googleapis.com/css2?family=Open+Sans:ital,wght@0,300..800;1,300..800&display=swap" rel="stylesheet">
        <style>
          * {
            margin: 0;
            padding: 0;
            box-sizing: inherit;
          }
      
          body {
            font-family: 'Open Sans', Arial, Helvetica, sans-serif;
            font-size: 16px;
            color: #333333;
            line-height: 130%;
            box-sizing: border-box;
          }
      
          table {
            border-collapse: collapse;
          }
        </style>
        <title>email-template-market-intelligence</title>
      </head>
      <body>
        <table
          border="0"
          cellspacing="0"
          cellpadding="0"
          role="presentation"
          style="width: 100%; background-color: #ffffff"
        >
          <tbody>
            <tr>
              <td style="margin: 0; padding: 20px 0" align="center">
                <table
                  border="0"
                  cellspacing="0"
                  cellpadding="0"
                  role="presentation"
                  style="width: 900px; margin: 0 auto; background-color: #ffffff; border-color: #dddddd;"
                  align="center"
                >
                  <tbody>
                    <tr>
                      <td
                        align="left"
                        style="font-family: 'Open Sans', Arial, Helvetica, sans-serif; margin: 0; padding: 0;"
                      >
                        <strong>Stay Ahead in HVAC/R: Your Bi-Weekly Digest of Latest Industry Trends & Insights</strong>
                      </td>
                    </tr>
                    <tr>
                      <td
                        style="margin: 0; padding: 15px 0; font-family:'Open Sans', Arial, Helvetica, sans-serif; font-size: 12px;"
                      >
                        <em>
                          Have you noticed some relevant events or HVAC/R news items that have not been covered? Please do write in to
                          <a href="mailto:lennoxmarketintelligence@lennox.com" style="color: #c60c35">LennoxMarketIntelligence@lennox.com</a>
                          with any feedback, questions or comments. Please also do write back if you would like to include any further members on the distribution list.
                        </em>
                      </td>
                    </tr>
                    `;

  // Looping through JSON
  jsonData.forEach((item) => {
    let organisationName = "";
    let descriptionListItems = "";
    let articleTitle = "";

    if (item.category === "Competitor") {
      item.articles.forEach((competitorArticle) => {
        organisationName = `<tr><td style="margin: 0; padding: 10px 0 0 0"><p style="font-size:14px; font-family:'Open Sans', Arial, Helvetica, sans-serif; line-height:130%">${competitorArticle.organisation_name}</p></td></tr>`;

        const descriptionPoints =
          competitorArticle.article_description.split(/\r?\n/);

        descriptionPoints.forEach((item) => {
          descriptionListItems += `<li>${item}</li>`;
          console.log(descriptionListItems);
        });

        articleTitle += `
        <tr><td>${organisationName}</td></tr>
        
        <tr>
         <td style="margin: 0; padding: 10px 0 0 20px">
        <p style="font-size:14px; font-family:'Open Sans', Arial, Helvetica, sans-serif; line-height:130%"><a href="${competitorArticle.article_link}" style="color:#c60c35" target="_blank">${competitorArticle.article_title}</a>
        &nbsp;&nbsp;${competitorArticle.article_date}
        </p>
        <ul style="padding: 10px 0 20px 20px; font-size:14px; font-family:'Open Sans', Arial, Helvetica, sans-serif; line-height:130%">
        ${descriptionListItems}
        </ul>
        </td>
        </tr>`;

        descriptionListItems = "";
      });
    } else {
      if (item.organisation_name && item.organisation_name.trim() !== "") {
        organisationName = `<tr><td style="margin: 0; padding: 10px 0 0 0"><p style="font-size:14px; font-family:'Open Sans', Arial, Helvetica, sans-serif; line-height:130%">${item.organisation_name}</p></td></tr>`;
      }

      if (item.article_description) {
        const descriptionPoints = item.article_description.split(/\r?\n/);
        descriptionPoints.forEach((point) => {
          descriptionListItems += `<li>${point}</li>`;
          console.log(descriptionListItems);
        });
      }

      if (item.article_title) {
        articleTitle += `
          <tr>
          <td style="margin: 0; padding: 10px 0 0 20px">
            <p style="font-size:14px; font-family:'Open Sans', Arial, Helvetica, sans-serif; line-height:130%"><a href="${item.article_link}" style="color:#c60c35" target="_blank">${item.article_title}</a>
            &nbsp;&nbsp;${item.article_date}
            <ul style="padding: 10px 0 20px 20px; font-size:14px; font-family:'Open Sans', Arial, Helvetica, sans-serif; line-height:130%">
              ${descriptionListItems}
            </ul></p>
          </td>
        </tr>`;
      }
    }

    emailTemplate += `
                    <tr>
                        <td style="margin: 0"><h4>${item.category}</h4></td>
                    </tr>
                    ${articleTitle}
          `;
  });

  emailTemplate += `
                    </tbody>
                  </table>
                </td>
              </tr>
            </tbody>
          </table>
        </body>
      </html>`;

  return emailTemplate;
}

function displayTemplate(emailTemplate) {
  document.getElementById("output").innerHTML = emailTemplate;
}

function downloadTemplate(emailTemplate) {
  document.getElementById("download").addEventListener("click", function () {
    const blob = new Blob([emailTemplate], { type: "text/html" });
    const url = window.URL.createObjectURL(blob);
    const downloadLink = document.createElement("a");
    downloadLink.href = url;
    downloadLink.download = "html_from_excel.html";
    downloadLink.click();
    window.URL.revokeObjectURL(url);
  });
}

document.getElementById("input").addEventListener("change", handleFileChange);
