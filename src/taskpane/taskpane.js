/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.onReady(info => {
    if (info.host === Office.HostType.Excel) {
        var timeouts = []
        document.getElementById('clear').onclick = clearOnlclickHandler
        document.getElementById('coinTable').onclick = coinTableOnlclickHandler
        document.getElementById('testRefeshRate').onclick = () => {
            timeouts.push(setInterval(testRefeshRateOnlclickHandler, 1000))
        }
        document.getElementById('stopTestRefeshRate').onclick = () => {
            for (var i = 0; i < timeouts.length; i++) {
                clearTimeout(timeouts[i])
            }
        }
    }
})

/*
1. fetch binance API
2. convert JSON to array
3. reutrn array
*/
const fetchBinanceAPI = async () => {
    console.log('fetch api')
    var proxyUrl = 'https://cors-anywhere.herokuapp.com/'
    var targetUrl = 'https://api.binance.com/api/v1/ticker/24hr'
    let dataJson
    let res = await fetch(proxyUrl + targetUrl)
    if (res.ok) {
        dataJson = await res.json()
        /*
        dataText = Object.keys(dataJson).map(function(_) {
            return dataJson[_]
        })
        */

        //json to 2d array
        var i = 0
        var result = []

        while (i < dataJson.length) {
            result.push(
                Object.keys(dataJson[i]).map(function(_) {
                    return dataJson[i][_]
                })
            )
            i++
        }
    }
    return result
}

/*
clear the entire worksheet
*/
function clearWorksheet(context) {
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet()
    var usedRange = currentWorksheet.getRange()
    usedRange.clear()
}

/*
1. init table
2. fetch API and get data
3. add data to table
*/
const createCoinTable = async context => {
    console.log('create table')
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet()
    var expensesTable = currentWorksheet.tables.add('A1:U1', true /*hasHeaders*/)
    expensesTable.name = 'ExpensesTable'

    expensesTable.getHeaderRowRange().values = [
        [
            'symbol',
            'priceChange',
            'priceChangePercent',
            'weightedAvgPrice',
            'prevClosePrice',
            'lastPrice',
            'lastQty',
            'bidPrice',
            'bidQty',
            'askPrice',
            'askQty',
            'openPrice',
            'highPrice',
            'lowPrice',
            'volume',
            'quoteVolume',
            'openTime',
            'closeTime',
            'firstId',
            'lastId',
            'count',
        ],
    ]

    await fetchBinanceAPI().then(data => {
        console.log('finish fetch')
        try {
            //console.log(data)
            expensesTable.rows.add(null /*add at the end*/, data)
        } catch (e) {
            console.log('error in add')
        }
        //expensesTable.columns.getItemAt(3).getRange().numberFormat = [['€#,##0.00']]
        expensesTable.getRange().format.autofitColumns()
        expensesTable.getRange().format.autofitRows()
    })
}

function addRandomNumber(context) {
    var num = Math.random() * 1000 + 1
}

function testWorksheetRefreshRate(context) {
    console.log('test worksheet refresh rate')
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet()
    var string1 = 'A'
    var string2 = '1'
    var string3 = String.fromCharCode(string1.charCodeAt())
    for (var i = 0; i < 10; i++) {
        for (var j = 1; j <= 10; j++) {
            var finalString = String.fromCharCode(string1.charCodeAt() + i) + j.toString()
            console.log(finalString)
            var range = currentWorksheet.getRange(finalString)
            range.values = Math.random() * 1000 + 1
        }
    }
}
export async function clearOnlclickHandler() {
    try {
        await Excel.run(async context => {
            /**
             * Insert your Excel code here
             */
            clearWorksheet(context)

            await context.sync()
        })
    } catch (error) {
        console.error(error)
    }
}

export async function coinTableOnlclickHandler() {
    try {
        await Excel.run(async context => {
            /**
             * Insert your Excel code here
             */
            clearWorksheet(context)
            await createCoinTable(context)

            await context.sync()
        })
    } catch (error) {
        console.error(error)
    }
}

export async function testRefeshRateOnlclickHandler() {
    try {
        await Excel.run(async context => {
            /**
             * Insert your Excel code here
             */
            clearWorksheet(context)
            testWorksheetRefreshRate(context)

            await context.sync()
        })
    } catch (error) {
        console.error(error)
    }
}
