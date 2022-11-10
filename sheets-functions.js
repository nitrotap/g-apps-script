
// todo add spreadsheet id
const spreadsheetId = '';

/** @OnlyCurrentDoc */
async function getDateCOI() {
    // todo add master sheet name
    let range = 'Sheet1';
    try {
        let write_array = []
        let new_array = []
        const result = await Sheets.Spreadsheets.Values.get(spreadsheetId, range);

        // loop through sheet, comparing spreadsheet date with today
        for (i = 1; i < result.values.length; i++) {
            // todo add column integer
            let col = 8;

            let d1 = new Date(result.values[i][col]);
            let d2 = new Date();

            // if date is over 1 year, then add to array
            if (d2 - d1 > 31556952000) {
                new_array.push([result.values[i][1], result.values[i][0], result.values[i][3], result.values[i][4]])
            }
            write_array.push(new_array);
        }

        // set up new sheet headers
        // todo row headers
        const rowAValues = [
            ['first name', 'last name', 'email', 'phone']
        ]

        // set up request for api call
        // todo sheet name
        let sheet_name = 'LastCOI';
        const request = {
            'valueInputOption': 'USER_ENTERED',
            'data': [
                {
                    'range': sheet_name + '!A1',
                    'majorDimension': 'ROWS',
                    'values': rowAValues
                },
                {
                    'range': sheet_name + '!A2',
                    'majorDimension': 'ROWS',
                    'values': new_array
                }
            ]
        };

        console.log(rowAValues)

        // get request to api
        try {
            const response = await Sheets.Spreadsheets.Values.batchUpdate(request, spreadsheetId);
            if (response) {
                Logger.log(response);
                return;
            }
            Logger.log('response null');
        } catch (e) {
            Logger.log('Failed with error %s', e.message);
        }
    } catch (e) {
        console.log(e)
    }
}

/** @OnlyCurrentDoc */
async function getBC() {
    // todo for sheet name
    let range = 'Sheet1';
    try {
        let write_array = []
        let new_array = []
        const result = await Sheets.Spreadsheets.Values.get(spreadsheetId, range);


        // loop through sheet, comparing spreadsheet date with today
        for (i = 1; i < result.values.length; i++) {
            // todo column integer
            let col = 7;

            let d1 = new Date(result.values[i][col]);
            let d2 = new Date();

            // if date is over 1 year, then add to array
            if (d2 - d1 > 31556952000) {
                new_array.push([result.values[i][1], result.values[i][0], result.values[i][3], result.values[i][4]])
            }
            write_array.push(new_array);
        }

        // set up new sheet headers
        const rowAValues = [
            ['first name', 'last name', 'email', 'phone']
        ]

        // set up request for api call
        // todo sheet name
        let sheet_name = 'BC';
        const request = {
            'valueInputOption': 'USER_ENTERED',
            'data': [
                {
                    'range': sheet_name + '!A1',
                    'majorDimension': 'ROWS',
                    'values': rowAValues
                },
                {
                    'range': sheet_name + '!A2',
                    'majorDimension': 'ROWS',
                    'values': new_array
                }
            ]
        };

        console.log(rowAValues)

        // get request to api
        try {
            const response = await Sheets.Spreadsheets.Values.batchUpdate(request, spreadsheetId);
            if (response) {
                Logger.log(response);
                return;
            }
            Logger.log('response null');
        } catch (e) {
            Logger.log('Failed with error %s', e.message);
        }
    } catch (e) {
        console.log(e)
    }
}
