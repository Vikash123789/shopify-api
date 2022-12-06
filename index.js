require("dotenv").config();
const express = require("express")
let axios = require("axios")
const app = express();
const { GoogleSpreadsheet } = require('google-spreadsheet');
const creds = require('./secret/extreme-storm-369107-c2750b32474a.json')
const _ = require('lodash');
const dayjs = require('dayjs');
const getSymbolFromCurrency = require('currency-symbol-map');

// load plugins
var customParseFormat = require('dayjs/plugin/customParseFormat');
const { last } = require("lodash");
dayjs.extend(customParseFormat);





var lastUpdated = {
    min: dayjs().subtract(65, 'minute').toISOString(),
    max: dayjs().toISOString()
}

let options = {
    method: 'GET',
    url: `https://${process.env.API_KEY}:${process.env.API_SECRET_KEY_WITH_TOKEN}@${process.env.STORE_NAME}/admin/api/${process.env.API_VERSION}/orders.json?status=any&updated_at_min=${lastUpdated.min}`
}




const prettyDate = $date => {
    return dayjs($date).format('DD MMM YYYY');
};


app.get('/', async (req, res) => {

let tsw = req.query.tsw
   // let tsw = req.headers.tsw
  

    if (tsw == process.env.SECRET_VALUE) {

        let wrapperFunction = async () =>{

     

        let removedRows = [];

        var currentTime = new Date();

        let result = await axios(options)
        let data = result.data.orders
        let jsonToCsv = []

        let jsonData = () => {

            for (let item of data) {
                for (let i = 0; i < item.line_items.length; i++) {
                    jsonToCsv.push(
                        {
                            "ID": item.name,
                            "Order Time (DD/MMM/YYYY)": prettyDate(item.processed_at),
                            "Customer Name": item.customer.default_address.name,
                            "Customer Email": item.customer.email,
                            "Currency": item.customer.currency,
                            "Price":  item.line_items[i].price  ,
                            "Total Discount" : item.line_items[i].total_discount ? item.line_items[i].total_discount : 0 ,
                            "Quantity": item.line_items[i].quantity,
                            "Item Name": item.line_items[i].title,
                            "Item SKU": item.line_items[i].sku,
                            "Variant Type": item.line_items[i].variant_title,
                            "Fulfillment Status": item.line_items[i].fulfillment_status ? item.fulfillment_status : 'Unfulfilled',
                            "Payment Status": (item.financial_status == "paid")? "Paid" : "Pending",
                            "Cancelled": (item.cancelled_at == null) ? "No" : "Yes",
                            "City": item.customer.default_address.city,
                            "Fullfiment Date": (item.fulfillments.length == 0) ? null : prettyDate(item.fulfillments[0].created_at),
                            "Delivery Vendor": (item.fulfillments.length == 0) ? null : item.fulfillments[0].tracking_company,
                            "Delivery Type": item.shipping_lines[0].title ?  item.shipping_lines[0].title : null  ,
                            "Delivered Date": (item.fulfillments.length == 0) ? null : (item.fulfillments[0].shipment_status == "delivered")? prettyDate(item.fulfillments[0].updated_at) : null,
                            "Fullfiment QTY": item.line_items[i].fulfillable_quantity,
                            "Remarks (Reason for cancellation/ delay)  ": item.cancel_reason,
                            "Transaction Type": item.gateway,
                            "Row Updated": "no",
                            "Row Updated Time": prettyDate(currentTime),
                        }
                    )
                }
            }

        }


        jsonData()

        const doc = new GoogleSpreadsheet('1UKHQJuZPgzWw6nhGn3p5XIw5aloQr1pVFnE1E46rW0c');
        await doc.useServiceAccountAuth(creds);
        await doc.loadInfo();
        console.log(doc.title);


        await doc.updateProperties({ title: 'Pengolin Order Data' });
        const sheet = doc.sheetsByIndex[0];


        const HEADERS = ["ID",
            "Order Time (DD/MMM/YYYY)",
            "Customer Name",
            "Customer Email",
            "Currency",
            "Price",
            "Total Discount",
            "Quantity",
            "Item Name",
            "Item SKU",
            "Variant Type",
            "Fulfillment Status",
            "Payment Status",
            "Cancelled",
            "City",
            "Fullfiment Date",
            "Delivery Vendor",
            "Delivery Type",
            "Delivered Date",
            "Fullfiment QTY",
            "Remarks (Reason for cancellation/ delay)",
            "Transaction Type",
            "Row Updated",
            "Row Update Time",
        ]

        await sheet.setHeaderRow(HEADERS);


        const getRows = await sheet.getRows();

        let newData = []



        // Google spreadsheet rows
        let spreadsheetRows = [];
        let alreadyExists = [];
        console.log(getRows.length);

        if (getRows.length > 0) {
            let totalUpdatedRows = 0;
            for (let i = 0; i < getRows.length; i++) {
                let obj = {
                    'ID': getRows[i].ID,
                    'Item SKU': getRows[i]['Item SKU']
                };

                spreadsheetRows.push(obj);

                let findInCsv = _.find(jsonToCsv, obj);
                if (findInCsv) {
                    // this means we already found it
                    alreadyExists.push(findInCsv);

                    // Now update it accordingly
                    // getRows[i]["Payment Status"] = findInCsv["Payment Status"];
                    // getRows[i]["Updated"] = "yes";

                    for (var key in findInCsv) {
                        getRows[i][key] = findInCsv[key];
                    }
                    getRows[i]['Row Updated'] = 'Yes';
                    getRows[i]['Row Update Time'] = prettyDate(currentTime);



                    // Update everything maybe, 

                    // Once row is found remove that from the jsonToCSV
                    let removedObj = _.remove(jsonToCsv, function (o) {
                        return o.ID == getRows[i].ID && o['Item SKU'] == getRows[i]['Item SKU'];
                    });
                    removedRows.push(removedObj);



                    // save the data
                    await getRows[i].save();
                    totalUpdatedRows++;



                    console.log('Found in CSV: ', findInCsv);

                } else {
                    // didn't find a matching row.
                }

                console.log(`Completed updating found rows. Total updated rows: ${totalUpdatedRows}`);
            }
        } else {
            // do nothing here..
        }

        newData = jsonToCsv;
        console.log(newData)
        await sheet.addRows(newData)

    

    
    let status = {
        success: 'ok',
        data: {
                foundRows: alreadyExists,
                addedRows: newData,
                removedRows: removedRows
            
            }
        }
        
        return res.send(status)
    }
        
        wrapperFunction()

    } else {

        return res.send({ status: "secret value not match" })
    }
    // return status
})




app.listen(process.env.PORT, function () {
    console.log('Express app running on port ' + (process.env.PORT))
});
