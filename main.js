const Excel = require('exceljs');

const convert = require('xml-js');
const fs = require('fs');

const colorSwapsXML = fs.readFileSync('ColorSwaps.xml', 'utf8');

var options = { ignoreComment: true, alwaysChildren: true };
const result = convert.xml2js(colorSwapsXML, options);
// console.log(result.elements[0].elements[0]);

const workbook = new Excel.Workbook();
workbook.creator = 'Alfie (twitter/AlfieBH_) • Corehalla.com';

const toHexString = (hex) => {
    return hex.replace('0x', '');
}

for (let i = 1; i < result.elements[0].elements.length; i++) {
    const colorScheme = result.elements[0].elements[i];
    if (i === 1) {
        console.log(colorScheme.elements[0]);
    }
    const body2Acc = colorScheme.elements.find(x => x.name === 'IndicatorColor');
    if (body2Acc) {
        const worksheet = workbook.addWorksheet(colorScheme.attributes.ColorSchemeName, {
            properties: {
                tabColor: toHexString(body2Acc.elements[0].text),
                defaultColumnWidth: 300,
                defaultRowHeight: 25
            }
        })

        worksheet.mergeCells('A1:E1');
        worksheet.getCell('A1').value = colorScheme.attributes.ColorSchemeName;
        worksheet.getCell('A1').font = {
            size: 20,
            bold: true
        };

        for (let j = 0; j < colorScheme.elements.length; j++) {
            const attribute = colorScheme.elements[j];

            worksheet.mergeCells(`A${j + 3}:B${j + 3}`);
            
            const isColor = attribute.elements[0].text.startsWith('0x');
            if (isColor) {
                worksheet.getCell(`A${j + 3}`).value = attribute.name.replace('_Swap', '');
                const hexColor = toHexString(attribute.elements[0].text);
                worksheet.getCell(`C${j + 3}`).fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: hexColor },
                    bgColor: { argb: hexColor }
                };
                worksheet.getCell(`D${j + 3}`).value = attribute.elements[0].text;
                worksheet.getCell(`E${j + 3}`).value = `#${hexColor}`;
            }
            else {
                worksheet.getCell(`A${j + 3}`).value = attribute.name;
                worksheet.mergeCells(`C${j + 3}:E${j + 3}`);
                worksheet.getCell(`C${j + 3}`).value = attribute.elements[0].text;
            }
        }

        worksheet.columns.forEach(column => {
            column.width = 20;
        })

    }
}

workbook.xlsx.writeFile('ColorSwaps.xlsx')
    .then(_ => console.log('Done!'))
    .catch(console.error)