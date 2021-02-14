const nodemailer = require("nodemailer");
const excelToJson = require("convert-excel-to-json");

// sourceFile = excel sheet, put the absolute or relative path to the excel file there
const data = excelToJson({
  sourceFile: "./Book.xlsx",
});

const sheets = Object.keys(data);

let transporter = nodemailer.createTransport({
  service: "gmail",
  auth: {
    user: "EMAIL",
    pass: "PASSW",
  },
});

// If platform is linux, remove leading "file:\\" from file path
if (process.platform === "linux" && data[sheets[0]][0]["B"].slice(0, 7) === "file://") {
  for (let i = 0; i < data[sheets[0]].length; i++) {
    data[sheets[0]][i]["B"] = data[sheets[0]][i]["B"].slice(7);
  }
}

for (let k = 0; k < sheets.length; k++) {
  for (let i = 0; i < data[sheets[k]].length; i++) {
    let mailOptions = {
      from: "SenderMail@gmail.com",
      to: data[sheets[k]][i]["A"],
      subject: "Testing mail",
      text: "Here's your file:",
      attachments: [
        {
          path: data[sheets[k]][i]["B"],
        },
      ],
    };

    transporter.sendMail(mailOptions, (e) => {
      if (e) {
        console.log("Error occurred:");
        console.log(e);
      }
    });
  }
}
