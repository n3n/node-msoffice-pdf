var mso_pdf = require("./lib");
var uuid = require("node-uuid");

module.exports = function excelToPdf(input) {
  return new Promise((resolve, reject) => {
    mso_pdf(null, function(error, office) {
      if (error) {
        console.log("Failed to init");
        return;
      }
      office.excel(
        {
          input: input,
          output: `output-${uuid.v4()}.pdf`
        },
        function(error, pdf) {
          if (error) {
            console.log("Excel: Failed to convert", error);
            reject(error);
          } else {
            console.log("Converted to: " + pdf);
            resolve(pdf);
          }
        }
      );

      office.close(null, function() {
        console.log("Office finished & closed, ");
      });
    });
  });
};
