require('dotenv/config')

const express = require('express')
const fileUpload = require('express-fileupload')
const morgan = require('morgan')
const Excel = require('exceljs')
const { data } = require('./data')

const app = express()

app.use(fileUpload({debug: false}))
app.use(express.json())
app.use(morgan('dev'))

const PORT = process.env.PORT || 8080


app.get('/', (req, res) => {
    res.sendFile(__dirname +  '/index.html')
})

app.post('/exceljs', async (req, res) => {
    const serviceId = req.body.id
    if (serviceId === undefined) {
        return res.status(400).json({'error': 'Service ID is required'})
    }
    const file = req.files.file
    if (file === null) {
        return res.status(406).json({'error': 'No files'})
    }


    allowedMimeTypes = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    ]
    if (!allowedMimeTypes.includes(file.mimetype)) {
        return res.status(415).json({'error': 'File must be .xlsx'})
    }
    
    if (file.size > 10 * 1024 * 1024) {
        return res.status(413).json({'error': 'File must have less than 10MB'})     
    }

    function getCellValueFromDB(cell, serviceId) {
        tag = cell.value
        tag = tag.substring(2, tag.length - 2)
        value = data[serviceId][tag]
        return value === undefined ? cell.value : value
    }
    
    const workbook = new Excel.Workbook();
    await workbook.xlsx.load(file.data);

    const worksheet = workbook.getWorksheet(1)

    worksheet.eachRow(function (row) {
        row.eachCell(function (cell) {
            if ( cell.value === null) {
                return
            }
            if (typeof cell.value === 'string') {
                
                if (cell.value.startsWith('<%') && cell.value.endsWith('%>')) {
                    cell.value = getCellValueFromDB(cell, serviceId)
                }
            }
            if (typeof cell.value === 'object') {
                
            }
        })
    })

    buffer = await workbook.xlsx.writeBuffer(tmpPath + 'response.xlsx');
    res.writeHead(200, {
        'Content-Type': 'application/pdfapplication/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename=Nota de Serviço.xlsx',
        'Content-Length': buffer.length
    });
    return res.end(buffer) 
})

app.listen(PORT, (req, res) => {
    console.log('Server listening on port ' + PORT)
})