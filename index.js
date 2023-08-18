const puppeteer = require('puppeteer');
const fs = require('fs');
const XLSX = require('xlsx');
const stringSimilarity = require('string-similarity');

//comparing strings to check for 90% similarilty
function areStringsSimilar(str1, str2) {
  const similarity = stringSimilarity.compareTwoStrings(str1, str2);
  return similarity >= 0.9;
}

// The path to Excel file
const filePath = 'Input.xlsx';

let jsonData=[];

// Read the Excel file using the fs module
fs.readFile(filePath, (err, data) => {
  if (err) {
    console.error('Error reading Excel file:', err);
    return;
  }

  // Parse the Excel data using xlsx
  const workbook = XLSX.read(data, { type: 'buffer' });

  // Reading the 1st sheet
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  // Convert the sheet data to JSON format
  jsonData = XLSX.utils.sheet_to_json(sheet);

  // Print the JSON data
  //console.log(jsonData);
});

async function run() {
    const browser = await puppeteer.launch({ headless: true });
    const page = await browser.newPage();
    await page.goto('https://www.snapdeal.com/');

    //await page.screenshot({path:'snapshot.png'}) //Checking if the browser's working as expected

    for(const item of jsonData){
        const el = String(item.ISBN)
        const name = String(item['Book Title'])
        //feeding the ISBN in the search field
        await page.type('input.col-xs-20.searchformInput.keyword', el);
        
        //searching for the item
        await page.keyboard.press('Enter');
        await page.waitForNavigation();
        
        //checking if a result is obtained or not 
        const noResult = await page.$('span.alert-heading') || null;
        if(noResult){
            //filling the required values for no result
            item['Found']='No';
            item['Price']=item['URL']=item['Author']=item['Publisher']=item['Availability']='NA';
        }
        else{
            item['Found'] = 'Yes';
            //storing the prices, links and urls of the search results 
            const bookData = await page.$$eval('div.product-desc-rating', allData => allData.map(data => data.innerHTML))
            
            //converting the data into an array of usable objects
            const scrappedData = await Promise.all(bookData.map(async htmlString => {
                const tempPage = await browser.newPage();
                await tempPage.setContent(htmlString);
            
                const title = await tempPage.$eval('p.product-title', el => el.textContent);
                const link = await tempPage.$eval('a.dp-widget-link.noUdLine.hashAdded', el => el.href);
                const price = await tempPage.$eval('span.lfloat.product-price', el => parseInt(el.innerText.substring(4)));
            
                await tempPage.close();
            
                return { title, link, price };
            }));
            console.log(scrappedData);

            let min=Infinity,url='';
            for (const val of scrappedData) {
                //checking for 90% similarity in book title and result obtained
                if (areStringsSimilar(val['title'].substring(0, name.length).toLowerCase(), name.toLowerCase())) {
                    if (min > val.price) {
                        min = val.price;
                        url = val.link;  
                    }
                }
            }            
            //Checking if a valid result has been obtained
            if(url!=''){
                item['Author'] = await page.$eval('p.product-author-name',val => val.textContent.substring(3)) || 'NIL'
                //Going to the url with the min price,i.e, the book page
                await page.goto(url)                
                item['Price']=min;
                item['URL']=url;
                //scrapping the data from the webpage
                item['Publisher'] = await page.$eval('#id-tab-container > div > div.spec-section.expanded.highlightsTileContent > div.spec-body.p-keyfeatures > ul > li:nth-child(3) > span.h-content',val=>val.textContent.substring(10))
                if(await page.$('div#add-cart-button-id'))
                    item['Availability'] = 'In Stock'
                else
                    item['Availability'] = 'Out of Stock'
            }
            else{
                //filling the required values for no valid result
                item['Found']='No';
                item['Price']=item['URL']=item['Author']=item['Publisher']=item['Availability']='NA';
            }
        }

        //await page.screenshot({ path: `search_res${item.No}.png` }); //screenshot for debugging 
        //selecting the value in the input field 
        await page.click('input.col-xs-20.searchformInput.keyword', { clickCount: 3 });
    }
    //console.log(jsonData);

    // Writing jsonData to Output file
    const ws = XLSX.utils.json_to_sheet(jsonData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    const excelFilePath = 'Output.xlsx';
    XLSX.writeFile(wb, excelFilePath, { bookType: 'xlsx', type: 'binary' });

    console.log(`Data has been written to ${excelFilePath}`);
    await browser.close();
}

run();