const pdf = require('html-pdf')

function generatePdfHtml(res, data) {

    const html = `
    <div>
        <h1>${data.nome}</h1>
        <h1>${data.email}</h1>
    </div>`
        
    const options = {
        type: 'pdf',
        format: 'A4',
        orientation: 'portrait'
    }

    pdf.create(html, options).toBuffer((err, buffer) => {
        if(err) return err
        console.log(buffer)
        res.writeHead(200, {
            'Content-Type': 'application/pdf',
            'Content-Disposition': 'attachment; filename=teste.pdf',
            'Content-Length': buffer.length
          });
        res.end(buffer)             
    })
}
module.exports = generatePdfHtml;

/* 
`
<!DOCTYPE html>
<html>
  <head>
    <style>
      ${css}
    </style>
  </head>
  <body id=report-pdf>
    ${html}
  </body>
</html>
`
*/