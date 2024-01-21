#!/usr/bin/env node
require('colors')
const fs = require('fs')
const path = require('path')
const { program } = require('commander');
const winax = require('winax');

async function main() {
    const option = init()
    var inputFile = path.resolve(option.inputFile)
    var outputDir = path.resolve(option.outputDir)
    console.log(`Output Dir: ${outputDir}`);

    if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir, { recursive: true })
    }

    let Excel = null;
    let workBook = null;
    try {
        Excel = new winax.Object('Excel.Application');
        workBook = Excel.Workbooks.Open(inputFile);
        const sheets = workBook.Worksheets;

        if (!option.split) {
            var outputFile = path.join(outputDir, `${workBook.Name}.pdf`)
            workBook.ExportAsFixedFormat(0, outputFile);

        } else {

            for (let i = 1; i <= sheets.Count; i++) {
                const sheet = sheets.Item(i);
                var outputFile = path.join(outputDir, `${sheet.Name}.pdf`)
                if (fs.existsSync(outputFile)) {
                    fs.rmSync(outputFile)
                }
                
                sheet.ExportAsFixedFormat(0, outputFile);
                console.log(`Export ${i}/${sheets.Count}: ${sheet.Name}`);
            }
        }

    } catch (e) {
        console.log(`${e}`.red);
        if (workBook != null) {
            workBook.Close(false);
        }
        if (Excel != null) {
            Excel.Quit()
        }
    }

}

main().then(() => {
    console.log("done!".green);
    process.exit(0)
})


/**
 * @returns {{config, table}}
 */
function init() {
    program.name('excel2pdf')
        .description("Convert Excel To PDF(s) by COM+ Component")
        .version('0.0.1')

    program.option('-i, --input-file [input.xlsx]', "Input Excel File", 'input.xlsx')
    program.option('-o, --output-dir [output]', "Output directory", 'output')
    program.option('-s, --split', "Split single worksheet")

    program.parse()
    const option = program.opts()
    return option
}




