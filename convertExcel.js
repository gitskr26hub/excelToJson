/**
 * Convert all Excel files in a folder into one JSON file
 * Clean, simple Node.js code from scratch
 */

import fs from "fs/promises";
import path from "path";
import xlsx from "xlsx";

async function convertExcelFolderToJson() {
  try {
    const folderPath = "./nitin";   // <-- Your folder name
    const outputPath = "./nitin.json"; // <-- Output JSON file

    // Resolve full absolute paths
    const absFolder = path.resolve(folderPath);
    const absOutput = path.resolve(outputPath);

    // Read all files inside folder
    let files;
    try {
      files = await fs.readdir(absFolder);
    } catch (err) {
      console.error("❌ Unable to read folder:", err.message);
      return;
    }

    if (files.length === 0) {
      console.error("❌ Folder is empty. No files to process.");
      return;
    }

    // Filter only Excel files
    const excelFiles = files.filter(file =>
      file.endsWith(".xlsx") || file.endsWith(".xls")
    );

    if (excelFiles.length === 0) {
      console.error("❌ No Excel files found in folder.");
      return;
    }

    let finalJsonData = [];

    // Process each excel file
    for (const fileName of excelFiles) {
      const filePath = path.join(absFolder, fileName);

      try {
        const workbook = xlsx.readFile(filePath);
        const firstSheetName = workbook.SheetNames[0];
        const firstSheet = workbook.Sheets[firstSheetName];

        const jsonData = xlsx.utils.sheet_to_json(firstSheet);

        finalJsonData.push(...jsonData);

        console.log(`✔ Converted: ${fileName}`);
      } catch (err) {
        console.error(`❌ Failed to convert ${fileName}:`, err.message);
      }
    }

    // Save final JSON
    try {
      await fs.writeFile(absOutput, JSON.stringify(finalJsonData, null, 2));
      console.log(`\n🎉 JSON generated successfully at: ${absOutput}`);
    } catch (err) {
      console.error("❌ Error writing JSON file:", err.message);
    }

  } catch (error) {
    console.error("🔥 Unexpected Error:", error.message);
  }
}

// Run the function
convertExcelFolderToJson();
