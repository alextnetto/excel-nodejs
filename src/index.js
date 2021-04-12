require('dotenv/config')

const express = require('express')
const fileUpload = require('express-fileupload')
const path = require('path')
var xlsx = require('node-xlsx')
const morgan = require('morgan')
const { data } = require('./data')
const generatePdfHtml = require('./utils/pdfHtml')
const Excel = require('exceljs')

const app = express()

app.use(fileUpload({debug: false}))
app.use(express.json())
app.use(morgan('dev'))

const PORT = process.env.PORT || 8080


app.get('/', (req, res) => {
    res.sendFile(__dirname +  '/index.html')
})


app.post('/excel/upload', async (req, res) => {
    console.log(req.files)

    if (req.files.file) {
        file = req.files.file
        const tmpPath = path.resolve(__dirname, '..', 'tmp', 'uploads') + '/'


        allowedMimeTypes = [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        ]

        if (file.size > 2 * 1024 * 1024) {
            return res.status(413).send({'error': 'File must be less than 10MB'})     
        }
        if (!allowedMimeTypes.includes(file.mimetype)) {
            return res.status(415).send({'error': 'File type not suported'})
        }

        //Save file in tmp folder
        file.mv(tmpPath + file.name)

        //Parse excel buffer to object
        var xlsxData = xlsx.parse(file.data)[0].data

        // TODO: Get user data from DB with id
        userData = data[req.body.id]
        finalResponse = {}

        newXlsxData = xlsxData.map((row) => {
            return row.map((cell) => {
                cell = String(cell)
                if (cell.startsWith('<%') && cell.endsWith('%>')) {
                    cellTag = cell.substring(2, cell.length - 2)
                    cellValue = userData[cellTag]
                    finalResponse[cellTag] = cellValue
                    return cellValue
                }
                return cell
            })
        })

        generatePdfHtml(res, finalResponse)
        // pdfBuffer = await generatePdfHtml()
        // return res.end(pdfBuffer)
        // return res.download(tmpPath + 'Finalizado - 4642.pdf')
    } else {
        return res.status(406).send({'error': 'No files'})
    }
})

app.post('/exceljs', async (req, res) => {
    const serviceId = req.body.id
    if (serviceId === undefined) {
        return res.status(400).json({'error': 'Service ID is required'})
    }

    function getCellValueFromDB(cell, serviceId) {
        tag = cell.value
        tag = tag.substring(2, tag.length - 2)
        console.log(data[serviceId], tag)
        value = data[serviceId][tag]
        console.log(value === undefined ? cell.value : value)
        return value === undefined ? cell.value : value
    }
    
    const tmpPath = path.resolve(__dirname, '..', 'tmp', 'uploads') + '/'

    const workbook = new Excel.Workbook();
    await workbook.xlsx.load(req.files.file.data);
    data1 = []
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
    res.end(buffer) 

})

app.listen(PORT, (req, res) => {
    console.log('Server listening on port ' + PORT)
})