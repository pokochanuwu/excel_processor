# excel_processor


Excel Read/Write module


```js
const excel = require('./module')
const filename = 'ee_data.xlsx'
const sheetname = 'data'
const area="A1:D4"

main()
async function main() {
    //read
    let FetchedData = await excel.read(area, sheetname, filename)
    //write
    await excel.write(FetchedData, area, sheetname, filename)
}
``` 
