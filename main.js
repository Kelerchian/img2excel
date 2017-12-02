const Jimp = require("jimp")
const Excel = require("exceljs")
const fs = require("fs")
const path = require("path")

const inputFolder = path.resolve(__dirname,"playground","input")
const outputFolder = path.resolve(__dirname,"playground","output")

async function main(){

  if(!fs.existsSync(outputFolder)) fs.mkdirSync(outputFolder)
  if(!fs.existsSync(inputFolder)) fs.mkdirSync(inputFolder)

  var allowedExtensions = ["png","jpeg","jpg","bmp"]

  var filenames = fs.readdirSync(inputFolder)
    .filter(function(filename){
      return allowedExtensions.reduce(function(valid, allowedExtension){
        if(filename.endsWith(allowedExtension)) valid = true
        return valid
      },false)
    })

  if(filenames.length <= 0){
    console.log("No image file found")
    return
  }

  //load image
  var image = await new Promise(function(resolve, reject){
    Jimp.read( path.resolve(inputFolder, filenames[0]), function(err, image){
      if(err){ reject.err }
      resolve(image)
    })
  })

  var maxSize = 200
  if(image.bitmap.width > maxSize || image.bitmap.height > maxSize){
    image.scaleToFit(maxSize,maxSize,Jimp.RESIZE_BICUBIC)
  }

  var width = image.bitmap.width
  var height = image.bitmap.height
  var workbook = new Excel.Workbook()
  var worksheet = workbook.addWorksheet('image', {
    properties:{
      defaultRowHeight: 1,
      defaultColumnWidth: 1
    }
  })

  for(var x = 1; x<=width; x++){
    worksheet.getColumn(x).width = 1
  }
  for(var y = 1; y<=height; y++){
    worksheet.getRow(y).height = 1
  }
  image.scan(0,0, width, height, function(x,y,idx){
    var cell = worksheet.getRow(y+1).getCell(x+1)
    var r = this.bitmap.data[idx+0]
    var g = this.bitmap.data[idx+1]
    var b = this.bitmap.data[idx+2]
    var a = this.bitmap.data[idx+3]

    cell.fill = {
      type:"pattern",
      pattern:"solid",
      fgColor: {argb: `${a.toString(16)}${r.toString(16)}${g.toString(16)}${b.toString(16)}` },
      bgColor: {argb: `${a.toString(16)}${r.toString(16)}${g.toString(16)}${b.toString(16)}` }
    }
  })
  for(var y = 1; y<=height; y++){
    worksheet.getRow(y).commit()
  }


  // for(var y = 1; y<=height; y++){
  //   var row = worksheet.getRow(y)
  //   // row.height = 1
  //   for(var x = 1; x<=width; x++){
  //     // if(y == 1) worksheet.getColumn(x).width = 1
  //
  //     var cell = row.getCell(x)
  //     var color = Jimp.intToRGBA(image.getPixelColor(x,y))
  //
  //     cell.fill = {
  //       type: 'pattern',
  //       pattern: 'solid',
  //       fgColor: {argb: 'FF'+color.r.toString(16)+color.g.toString(16)+color.b.toString(16) },
  //       bgColor: {argb: 'FF'+color.r.toString(16)+color.g.toString(16)+color.b.toString(16) }
  //     }
  //   }
  //   row.commit()
  // }

  await workbook.xlsx.writeFile(path.resolve(outputFolder,"output.xlsx"))

  return "success"
}
main()
  .then(function(data){
    console.log(data || "")
  })
  .catch(function(err){
    console.error(err)
  })
