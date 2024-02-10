const xlsxPopulate = require('xlsx-populate');
const prompt = require('prompt');
const fs = require('fs');

// Define schema for user input
const schema = {
    properties: {
        filePath: {
            description: 'Enter the file path of the Excel file:',
            required: true
        }
    }
};

// Start the prompt
prompt.start();

// Prompt user for file path
prompt.get(schema, async (err, result) => {
    if (err) {
        console.error("Error:", err);
        return;
    }

    // Check if the file exists
    if (!fs.existsSync(result.filePath)) {
        console.error("Error: The specified file does not exist.");
        return;
    }

    // Check if the file has the correct format
    try {
        await xlsxPopulate.fromFileAsync(result.filePath);
    } catch (error) {
        console.error("Error: Invalid Excel file format.");
        return;
    }

    try {
        // Read the Excel file
        const workbook = await xlsxPopulate.fromFileAsync(result.filePath);
        const sheet = workbook.sheet(0); // Get the first sheet

        // Add column headers for BonusPercentage and BonusAmount in the first row
        sheet.cell("C1").value("BonusPercentage");
        sheet.cell("D1").value("BonusAmount");

        // Get the used range of the sheet
        const usedRange = sheet.usedRange();
        const numRows = usedRange.endCell().rowNumber(); // Corrected: usedRange.endCell().rowNumber()

        // Loop through each row of employee data (starting from row 2)
        for (let i = 2; i <= numRows; i++) {
            const salaryCell = sheet.cell(`B${i}`); // Assuming salary is in column B (2nd column)
            const salary = salaryCell.value();

            let bonusPercentage, bonusAmount;

            // Calculate bonus based on salary
            if (salary < 50000) {
                bonusPercentage = 5;
            } else if (salary <= 100000) {
                bonusPercentage = 7;
            } else {
                bonusPercentage = 10;
            }

            // Calculate bonus amount
            bonusAmount = (bonusPercentage / 100) * salary;

            // Write bonus percentage and bonus amount to new columns
            const bonusPercentageCell = salaryCell.relativeCell(0, 1); // Next column
            const bonusAmountCell = salaryCell.relativeCell(0, 2); // Next column after bonus percentage
            bonusPercentageCell.value(bonusPercentage);
            bonusAmountCell.value(bonusAmount);
        }

        // Save the workbook to a new Excel file
        await workbook.toFileAsync("./employee_data_with_bonus.xlsx");
        console.log("Bonus information added and Excel file saved successfully.");
    } catch (error) {
        console.error("Error:", error.message);
    }
    
});
