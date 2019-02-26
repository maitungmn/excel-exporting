const express = require('express');
const bodyParser = require('body-parser');
const xl = require('excel4node');

const pause = require('./lib/pause');

process.setMaxListeners(Infinity);

let app = express();
app.use(bodyParser.urlencoded({extended: false}));
app.use(bodyParser.json());

app.use(function (req, res, next) {
    res.header("Access-Control-Allow-Origin", "*");
    res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept");
    next();
});

app.post('/export', function (req, res) {
    let headers = req.body.headers;
    let data = req.body.data;
    let channelList = req.body.channelList;
    let ws = [];
    let wb = new xl.Workbook();
    let myStyle = wb.createStyle({
        font: {
            bold: true,
            size: 14
        },
        alignment: {
            wrapText: true,
            horizontal: 'center',
            vertical: 'center'
        },
        border: {
            left: {
                style: 'thin',
                color: 'gray-50'
            },
            right: {
                style: 'thin',
                color: 'gray-50'
            },
            top: {
                style: 'thin',
                color: 'gray-50'
            },
            bottom: {
                style: 'thin',
                color: 'gray-50'
            },
            diagonal: {
                style: 'thin',
                color: 'gray-50'
            },
        },
        fill: {
            type: 'pattern',
            patternType: 'solid',
            bgColor: 'light yellow',
            fgColor: 'light yellow'
        }
    });

    for (let n = 0; n < channelList.length; n++) {
        ws[n] = wb.addWorksheet(channelList[n].name);

        for (let i = 1; i <= headers[n].length; i++) {
            ws[n].cell(1, i)
                .string(headers[n][i - 1].text)
                .style(myStyle);

            for (let j = 1; j <= data[n].length; j++) {
                for (let k = 1; k <= headers[n].length; k++) {
                    let X = '';
                    X = String(data[n][j - 1][headers[n][k - 1].value]);
                    ws[n].cell(j + 1, k).string(X).style({
                        alignment: {
                            horizontal: 'right'
                        }
                    });
                }
            }
        }

        wb.write('kpi-table.xlsx');

    }
    console.log('Done!')

    // res.send({
    //     headers: headers,
    //     data: data
    // })
});

app.set('port', process.env.PORT || 8080);
app.listen(app.get('port'), () => console.log('App is running on port ' + app.get('port')));
