const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs'); // Import exceljs to read the Excel file
const fs = require('fs');

// Function to read keywords from a specific sheet based on the day of the week
async function readKeywordsFromExcel(filePath, dayName) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath); // Read the Excel file

    // Check if the sheet for the given day exists
    const sheet = workbook.getWorksheet(dayName);
    if (!sheet) {
        console.log(`No sheet found for ${dayName}.`);
        return null;
    }

    const keywords = [];
    sheet.eachRow((row, rowNumber) => {
        // Assuming keywords are in the first column
        const keyword = row.getCell(1).value;
        
        // Skip the first row (the column header) and empty keywords
        if (keyword && keyword !== "Keyword") {
            keywords.push({ rowNumber, keyword });
        }
    });

    return { workbook, sheet, keywords }; // Return the workbook, sheet, and keywords
}

// Function to create a delay
function delay(time) {
    return new Promise(resolve => setTimeout(resolve, time));
}

// Function to scrape Google autocomplete suggestions
async function automateSearch(keyword) {
    const browser = await puppeteer.launch({ headless: false });
    const page = await browser.newPage();

    // Navigate to Google with English language preference
    await page.goto('https://www.google.com/?hl=en');

    // Type the search query
    await page.type('.gLFyf', keyword);

    // Add a slight delay to allow suggestions to load
    await delay(1000); // Wait for 1 second

    // Wait for suggestions to appear
    try {
        await page.waitForSelector('ul[role="listbox"]', { timeout: 10000 });
    } catch (error) {
        console.error('Error waiting for suggestions:', error);
        await browser.close();
        return { longest: '', shortest: '' };
    }

    // Get the suggestions
    const suggestions = await page.evaluate(() => {
        return Array.from(document.querySelectorAll('ul[role="listbox"] li')).map(suggestion => {
            return suggestion.innerText.split('\n')[0].trim(); // Take only the part before the newline and trim whitespace
        });
    });

    // Filter out empty or undefined suggestions
    const filteredSuggestions = suggestions.filter(s => s.length > 0);

    // Find the longest and shortest suggestions
    const longestSuggestion = filteredSuggestions.reduce((a, b) => a.length > b.length ? a : b, '');
    const shortestSuggestion = filteredSuggestions.reduce((a, b) => a.length < b.length ? a : b, filteredSuggestions[0]);

    // Close the browser
    await browser.close();

    return { longest: longestSuggestion, shortest: shortestSuggestion };
}

// Main function
async function main() {
    // Get the current day of the week (e.g., "Monday", "Tuesday", "Wednesday")
    const dayName = new Date().toLocaleString('en-US', { weekday: 'long' });

    const filePath = 'Excel.xlsx'; // Your original Excel file path

    // Read keywords from the sheet corresponding to the current day
    const { workbook, sheet, keywords } = await readKeywordsFromExcel(filePath, dayName);
    if (!keywords) return; // If no keywords found, exit

    // Create a new workbook for the updated data
    const updatedWorkbook = new ExcelJS.Workbook();
    const updatedSheet = updatedWorkbook.addWorksheet(dayName);

    // Add column headers to the new sheet
    updatedSheet.addRow(['Keyword', 'Longest', 'Shortest']);

    // Process the keywords and update the Excel sheet
    for (let { rowNumber, keyword } of keywords) {
        const { longest, shortest } = await automateSearch(keyword);

        // Write the keyword and its suggestions to the new updated sheet
        updatedSheet.addRow([keyword, longest, shortest]);

        console.log(`Processed keyword: ${keyword}`);
        console.log(`Longest: ${longest}`);
        console.log(`Shortest: ${shortest}`);
    }

    // Save the updated workbook with the name 'Updated_Excel.xlsx'
    await updatedWorkbook.xlsx.writeFile('Updated_Excel.xlsx');
    console.log(`Updated Excel file for ${dayName} created successfully!`);
}

main();
