import {api, LightningElement, track} from 'lwc';
import {loadScript} from "lightning/platformResourceLoader";
import workbook from "@salesforce/resourceUrl/sheetjs";
import LightningAlert from "lightning/alert";
export default class HandleExcel extends LightningElement {
    XLSX = {}; // excel Object
    version; // excel version
    columns = []; // excel header
    rows = []; // excel body
    async connectedCallback() {
        await loadScript(this, workbook +'/sheetjs.js') // load static resource
            .then(() => {
                this.XLSX = XLSX;
                console.log(this.XLSX);
                this.version = XLSX.version;
            })
            .catch((error) => {
                console.log('load error :: ' + error);
            });
    }

    get acceptedFormats() { // 파일 업로드 타입 : xlsx만 허용
        return ['.xlsx'];
    }
    uploadExcel (event) { // 엑셀 업로드
        let input = event.target;
        let type = event.target.files[0].type;
        let reader = new FileReader();
        reader.onload = () => {
            let data = reader.result;
            let workbook = this.XLSX.read(data, {type : 'binary'});
            workbook.SheetNames.forEach(sheetName => {
                let expData = this.XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
                if (this.validFileIsExcel(expData, type)) {
                    this.updateDatatable(expData);
                }
            });
        }
        reader.readAsBinaryString(input.files[0]);
    }
    downloadExcel () { // 엑셀 다운로드
        if (!(this.rows.length === 0)) {
            let sheetData = [];
            let keys = Object.keys(this.rows[0]);
            sheetData.push(keys);
            this.rows.forEach(data => {
                let body = [];
                for (let i = 0; i < keys.length; i++) {
                    let row = data[keys[i]];
                    body.push(row);
                }
                sheetData.push(body);
            });
            let wb =  this.XLSX.utils.book_new();
            let ws =  this.XLSX.utils.aoa_to_sheet(sheetData);
            this.XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
            this.XLSX.writeFile(wb, "data.xlsx");
        } else {
            this.handleAlert('다운가능한 데이터가 존재하지 않습니다. ', 'error', 'Error!').then((result) =>{
                console.log(result);
            });
        }

    }
    updateDatatable (expData) {
        let cols = [];
        let rowData = [];
        expData.forEach((data, idx) => {
            if (idx === 0) {
                let keys = Object.keys(data);
                for (let i = 0; i < keys.length; i++) {
                    let colAttribute = {
                        label : keys[i],
                        fieldName : keys[i],
                        type: 'text'
                    };
                    cols.push(colAttribute);
                }
            }
            else {
                rowData.push(data);
            }
        });
        this.columns = cols;
        this.rows = rowData;
    }

    validFileIsExcel (expData, type) {
        if (expData === null || expData === undefined || expData.length === 0) {
            this.handleAlert('엑셀 업로드에 실패했습니다.', 'error', 'Error!').then((result) =>{
                console.log(result);
            });

            return false;
        }
        else if (type !== 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
            this.handleAlert('엑셀 파일만(xlsx) 업로드 가능합니다.', 'error', 'Error!').then((result) =>{
                console.log(result);
            });

            return false;
        }

        return true;
    }

    async handleAlert (msg, thm, lbl) {
        await LightningAlert.open({
            message: msg,
            theme: thm,
            lbl: lbl
        });
    }

}