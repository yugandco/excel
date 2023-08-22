const express = require('express')
const xlsx = require('xlsx')
const multer = require('multer')
const ejs = require('ejs')
const path = require('path')
const bodyParser = require('body-parser')
const geolib = require('geolib')
const fs = require('fs')
const AdmZip = require('adm-zip')
const tokml = require('tokml')
// File Storage
const storage = multer.diskStorage({
    destination: './public/uploads/',
    filename: function(req, file, cb) {
        cb(null, file.originalname)
    }
})

const upload1 = multer({storage: storage}).single('ktPoints')

// Init Application
const app = express()

app.set('view engine', 'ejs')
app.use(express.static('./public'))
app.use(bodyParser.urlencoded({extended: false}))
app.use(bodyParser.json())


// Function to read files
const exportExcel = (infos, workSheetColumnNames, workSheetName, filePath) => {
    const workBook = xlsx.utils.book_new()

    const workSheet = xlsx.utils.json_to_sheet(infos, {header: workSheetColumnNames})
    xlsx.utils.book_append_sheet(workBook, workSheet, workSheetName)
    xlsx.writeFile(workBook, path.resolve(filePath))
}


// Get Home Page
app.get('/', (req, res) => {
    res.render('index')
})

// Upload File
app.post('/upload', (req, res) => {
    upload1(req, res, (err) => {
        if(err){
            res.render('index', {
                msg: err
            })
        } else {
            console.log(req.file)
            const ss = req.file.originalname
            const city = ss.substring(0, ss.indexOf("_"))
            
            const folderName = `./public/uploads/${city}`
            try {
                if(!fs.existsSync(folderName)){
                    fs.mkdirSync(folderName)
                }
            } catch (err) {
                console.error(err)
            }
    
            let output1 = req.body.textOne
            output1 = output1 === '' ? '' : JSON.parse(output1.trim())
            
            let output2 = req.body.textTwo 
            output2 = output2 === '' ? '' : JSON.parse(output2.trim())
            
            let output3 = req.body.textThree 
            output3 = output3 === '' ? '' : JSON.parse(output3.trim())
            
            let output4 = req.body.textFour 
            output4 = output4 === '' ? '' : JSON.parse(output4.trim())
            
            let output5 = req.body.textFive 
            output5 = output5 === '' ? '' : JSON.parse(output5.trim())
            
            let output6 = req.body.textSix 
            output6 = output6 === '' ? '' : JSON.parse(output6.trim())
            
            let output7 = req.body.textSeven 
            output7 = output7 === '' ? '' : JSON.parse(output7.trim())
            
            let output8 = req.body.textEight 
            output8 = output8 === '' ? '' : JSON.parse(output8.trim())
            
            let output9 = req.body.textNine 
            output9 = output9 === '' ? '' : JSON.parse(output9.trim())
            
            let output10 = req.body.textTen 
            output10 = output10 === '' ? '' : JSON.parse(output10.trim())
            
            
            const wb1 = xlsx.readFile(req.file.destination + req.file.originalname, {cellDates: true})
            const points = xlsx.utils.sheet_to_json(wb1.Sheets[wb1.SheetNames[0]])

            console.log(points[0])

            const points_polygon = [
                {
                    id: 1,
                    coor: output1
                },
                {
                    id: 2,
                    coor: output2
                },
                {
                    id: 3,
                    coor: output3
                },
                {
                    id: 4,
                    coor: output4
                },
                {
                    id: 5,
                    coor: output5
                },
                {
                    id: 6,
                    coor: output6
                },
                {
                    id: 7,
                    coor: output7
                },
                {
                    id: 8,
                    coor: output8
                },
                {
                    id: 9,
                    coor: output9
                },
                {
                    id: 10,
                    coor: output10
                },
              ]

              const result = []

            for(let polygon of points_polygon){
                console.log(polygon)
                for(let data of points){
                    let point = {
                        lat: data.lat,
                        lon: data.lon
                    }

                    
                    let res = geolib.isPointInPolygon(point, polygon.coor)

                    if(res === true){
                        result.push({    
                            ...data
                        })
                    }
                }

                let obj = {
                    type: "FeatureCollection",
                    features: [
                        polygon.coor
                    ]
                }

                let kml = tokml(obj)

                fs.writeFile(`./public/uploads/${city}/${city + "_" + polygon.id}.kml`, kml, (err) => {
                    if(err)throw err;
                    console.log('Polygon converted to kml')
                })

                
                
                exportExcel(
                    result,
                    Object.keys(points[0]),
                    'Sheet1',
                    `./public/uploads/${city}/${city + "_" + polygon.id}.xlsx`
                )

                result.length = 0

                console.log('Finished')

                if(polygon.id === 10){
                    // const zip = new AdmZip()
                    // const outputFile = `${city}.zip`
                    // zip.addLocalFile(`./public/uploads/${city}`)
                    // zip.writeZip(`./public/uploads/${outputFile}`)

                    //res.setHeader('Content-type','text/html')
                    //res.send('Все завершилось! Если есть еще полигоны, то вам нужно нажать кнопку назад и очистить выполненные полигоны и вставить новые)')
                    res.send(`

                            <p>Все завершилось! Если есть еще полигоны, то вам нужно нажать кнопку назад и очистить выполненные полигоны и вставить новые</p>    
                            <p>10 полигонов готовы! Ссылка для скачивания внизу)</p>
                            <a href="/public/uploads/${city}/${city+'_'+1}.xlsx" download>
                                <span>Download ${city + '_' + 1}</span>
                            </a>
                            <br>
                            <a href="/public/uploads/${city}/${city+'_'+2}.xlsx" download>
                                <span>Download ${city + '_' + 2}</span>
                            </a>
                            <br>
                            <a href="/public/uploads/${city}/${city+'_'+3}.xlsx" download>
                                <span>Download ${city + '_' + 3}</span>
                            </a>
                            <br>
                            <a href="/public/uploads/${city}/${city+'_'+4}.xlsx" download>
                                <span>Download ${city + '_' + 4}</span>
                            </a>
                            <br>
                            <a href="/public/uploads/${city}/${city+'_'+5}.xlsx" download>
                                <span>Download ${city + '_' + 5}</span>
                            </a>
                            <br>
                            <a href="/public/uploads/${city}/${city+'_'+6}.xlsx" download>
                                <span>Download ${city + '_' + 6}</span>
                            </a>
                            <br>
                            <a href="/public/uploads/${city}/${city+'_'+7}.xlsx" download>
                                <span>Download ${city + '_' + 7}</span>
                            </a>
                            <br>
                            <a href="/public/uploads/${city}/${city+'_'+8}.xlsx" download>
                                <span>Download ${city + '_' + 8}</span>
                            </a>
                            <br>
                            <a href="/public/uploads/${city}/${city+'_'+9}.xlsx" download>
                                <span>Download ${city + '_' + 9}</span>
                            </a>
                            <br>
                            <a href="/public/uploads/${city}/${city+'_'+10}.xlsx" download>
                                <span>Download ${city + '_' + 10}</span>
                            </a>
                    `)
                }
            }
        }
    })
})


const PORT = 5000
app.listen(PORT, () => {
    console.log(`Server runned ${PORT}`)
})

module.exports = app