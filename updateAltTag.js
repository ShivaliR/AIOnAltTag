const { google } = require('googleapis');
const keys = require('./sr-gsheet-access-8f1f5bd9790a.json');
const tectalicOpenai = require('@tectalic/openai').default;

const openaiApiKey = process.env.OPENAI_API_KEY || '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'; 
const openaiClient = tectalicOpenai(openaiApiKey);

const spreadsheetId1 = '1YHpQWFP78Nx2t34Cy-oKeGrs0lu_kcVpwggbE_LeJMc';
const sheetNumber = 1; // Sheet number to refer to in the workbook

const spreadsheetId2 = '17-x1OYChOELaJdwB6HIs5FJhL4oltWCAFJtaltGhZoc';

// Create a new JWT client using the key file downloaded from the Google Cloud Console
const client = new google.auth.JWT(
  keys.client_email,
  null,
  keys.private_key,
  ['https://www.googleapis.com/auth/spreadsheets']
);

// Authorize the client
client.authorize(async function (err, tokens) {
  if (err) {
    console.log(err);
    return;
  }

  // Example: Process images and update sheet
  try {
    await processImagesAndUpdateSheet();
  } catch (error) {
    console.error('Error processing images:', error.message);
  }
});

async function processImagesAndUpdateSheet() {
  const sheets = google.sheets({ version: 'v4', auth: client });

  // Reading image URLs and Alt Tags from the Google Sheet 1
  const response1 = await sheets.spreadsheets.get({
    spreadsheetId: spreadsheetId1,
  });

  const sheetName1 = response1.data.sheets[sheetNumber - 1].properties.title;
  console.log(`Using sheet 1: ${sheetName1}`);

  const range1 = `${sheetName1}!A1:ZZ`; // Assuming the data starts from cell A1 and goes to column ZZ

  const valuesResponse1 = await sheets.spreadsheets.values.get({
    spreadsheetId: spreadsheetId1,
    range: range1,
  });

  const sheetData1 = valuesResponse1.data.values || [];

  // Reading data from Google Sheet 2
  const response2 = await sheets.spreadsheets.get({
    spreadsheetId: spreadsheetId2,
  });

  const sheetName2 = response2.data.sheets[sheetNumber - 1].properties.title;
  console.log(`Using sheet 2: ${sheetName2}`);

  const range2 = `${sheetName2}!A1:ZZ`; // Assuming the data starts from cell A1 and goes to column ZZ

  const valuesResponse2 = await sheets.spreadsheets.values.get({
    spreadsheetId: spreadsheetId2,
    range: range2,
  });

  const sheetData2 = valuesResponse2.data.values || [];

  // Check if there is at least one row of data (excluding the header)
  if (sheetData1.length < 2 || sheetData2.length < 2) {
    console.log('No data found in one of the sheets.');
    return;
  }

  // Extract header and data rows from both sheets
  const [header1, ...imageRows] = sheetData1;
  const [header2, ...keywordRows] = sheetData2;

  // Create an array to store CSV rows
  const outputCsvRows = [['Image URL', 'Alt Tag']];

  // Match URLs and process data
  for (let i = 0; i < imageRows.length; i++) {
    const imageUrl = (imageRows[i][1] || '').trim();
    const websiteURL = (imageRows[i][0] || '').trim();

    // Find corresponding row in sheet 2 based on URL
    const matchingRow = keywordRows.find(row => row[0] === websiteURL);

    let keywords, language;

    if (matchingRow) {
      keywords = matchingRow[1];
      language = matchingRow[2];
    } else {
      keywords = '';
      language = '';
    }

    if (keywords) {
      console.log(imageUrl);
      console.log(websiteURL);
      console.log(keywords);
      console.log(language);
      console.log('-------');

      try {
        // Generate Alt Tag using GPT-4
        const generatedAltTag = await analyzeAndGenerateAltTag(imageUrl, keywords, language);

        // Update the Alt Tag in the sheet
        await updateAltTagInSheet(i, generatedAltTag, sheetName1);

        console.log(`Generated Alt tag for ${imageUrl}: ${generatedAltTag}`);
      } catch (error) {
        console.error(`Error processing image ${imageUrl}:`, error.message);
        // If there is an error, still add the image URL to the CSV with an empty alt tag
        outputCsvRows.push([imageUrl, '']);
      }
    }
  }

  console.log('Alt tags updated in the sheet.');
}

async function analyzeAndGenerateAltTag(imageUrl, keywords, language) {
  try {

    let response;

    if (keywords != '' && language != '') {
      response = await openaiClient.chatCompletions.create({
        model: 'gpt-4-vision-preview',
        messages: [
          {
            "role": "user",
            "content": [
              {
                "type": "text",
                "text": `Generate alt tag for the image url in 4-5 words that i can implement on the website as per this keyword ${keywords} in ${language} language and not inside inverted commas, and not mention that this is an image of.`,
              },
              {
                "type": "image_url",
                "image_url": {
                  "url": imageUrl,
                },
              },
            ],
          }
        ],
      });
    } else if (keywords != '') {

      response = await openaiClient.chatCompletions.create({
        model: 'gpt-4-vision-preview',
        messages: [
          {
            "role": "user",
            "content": [
              {
                "type": "text",
                "text": `Generate alt tag for the image url in 4-5 words that i can implement on the website as per this keyword ${keywords} and not inside inverted commas, and not mention that this is an image of.`,
              },
              {
                "type": "image_url",
                "image_url": {
                  "url": imageUrl,
                },
              },
            ],
          }
        ],
      });
    } else {
      response = await openaiClient.chatCompletions.create({
        model: 'gpt-4-vision-preview',
        messages: [
          {
            "role": "user",
            "content": [
              {
                "type": "text",
                "text": `Generate alt tag for the image url in 4-5 words that i can implement on the website and not inside inverted commas, and not mention that this is an image of.`,
              },
              {
                "type": "image_url",
                "image_url": {
                  "url": imageUrl,
                },
              },
            ],
          }
        ],
      });
    }

    return response.data.choices[0].message.content.trim();
  } catch (error) {
    throw new Error(`Failed to analyze image: ${error.message}`);
  }
}

async function updateAltTagInSheet(rowIndex, altTag, sheetName) {
  try {
    const sheets = google.sheets({ version: 'v4', auth: client });

    await sheets.spreadsheets.values.update({
      spreadsheetId: spreadsheetId1,
      range: `${sheetName}!C${rowIndex + 2}:C${rowIndex + 2}`,
      valueInputOption: 'RAW',
      resource: {
        values: [[altTag]],
      },
    });

    console.log(`Alt tag updated in the sheet at row ${rowIndex + 2}: ${altTag}`);
  } catch (error) {
    throw new Error(`Failed to update alt tag in sheet: ${error.message}`);
  }
}
