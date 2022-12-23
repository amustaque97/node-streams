import fetch from 'node-fetch';
import {
    Readable,
    pipeline,
    Transform
} from 'stream';
import exceljs from 'exceljs';
import util from 'util';

const URL = 'https://jsonplaceholder.typicode.com/posts/'


const cpipeline = util.promisify(pipeline);

class ExcelStream extends Transform {
    constructor(options = {}) {
        super({
            ...options,
            objectMode: true
        });
        this.i = 0;
        const o = {
            filename: "./test-file.xlsx",
            useStyles: true,
            useSharedStrings: true,
        };
        this.workbook = new exceljs.stream.xlsx.WorkbookWriter(o);
        this.workbook.creator = "ABCDEF";
        this.workbook.created = new Date();
        this.sheet = this.workbook.addWorksheet("Discussion Event");
        this.sheet.columns = [{
            header: "testColumn",
            width: 10,
            style: {
                font: {
                    bold: true,
                },
            },
        }, ];
    }

    async _transform(chunk, encoding, done) {
        try {
            this.i += 1;
            this.sheet.addRow([chunk]).commit();
        } catch (error) {
            console.log("Error while transforming data - commit failed");
        }
    }

    async _flush(cb) {
        try {
            this.sheet.commit();
            await this.workbook.commit();
        } catch (error) {
            console.log("Error cannot commit workbook");
        }
    }
}


async function run() {
    let i = 1;
    const output = new ExcelStream()
    while (true) {
        let u = URL + i
        const response = await fetch(u)
        const result = await response.json()
        if (result) {
            console.log(result)
            cpipeline(Readable.from([result]),
                async function* transform(src) {
                        for await (const chunk of src) {
                            yield JSON.stringify(chunk);
                        }
                    },
                    output)
        } else {
            break
        }


        i++
        if (i == 20) break

    }
}

run()
