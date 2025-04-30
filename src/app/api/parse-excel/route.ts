import { NextRequest, NextResponse } from 'next/server';
import * as XLSX from 'xlsx';

// Increase the body size limit for file uploads (50MB)
export const config = {
  api: {
    bodyParser: false,
    responseLimit: '50mb',
  },
};

interface ExtractedData {
  poSlNo: string | null;
  itemDescription: string | null;
  engineeringApprovalDate: string | null;
  procurementStartDate: string | null;
  deliveryAtSite: string | null;
}

interface FileResult {
  fileName: string;
  success: boolean;
  error?: string;
  data?: ExtractedData[];
  sheetName?: string;
}

// Type for the row data from Excel
interface ExcelRow {
  [key: string]: string | number | undefined;
}

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const files = formData.getAll('files') as File[];
    
    if (!files || files.length === 0) {
      return NextResponse.json(
        { error: 'No files uploaded' },
        { status: 400 }
      );
    }
    
    if (files.length > 20) {
      return NextResponse.json(
        { error: 'Maximum 20 files allowed' },
        { status: 400 }
      );
    }
    
    // Process each file
    const fileResults: FileResult[] = await Promise.all(
      files.map(async (file) => {
        try {
          // Validate file type
          const isValidType = 
            file.name.endsWith('.xls') || 
            file.name.endsWith('.xlsx');
            
          if (!isValidType) {
            return {
              fileName: file.name,
              success: false,
              error: 'Invalid file type. Only Excel files (.xls, .xlsx) are allowed',
            };
          }
          
          // Convert file to array buffer
          const buffer = await file.arrayBuffer();
          
          // Parse Excel file
          const workbook = XLSX.read(buffer, { type: 'array' });
          
          if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
            return {
              fileName: file.name,
              success: false,
              error: 'No sheets found in Excel file',
            };
          }
          
          // Try to find a sheet with our required columns
          let validSheetName: string | null = null;
          let headerRowIndex = -1;
          let subHeaderRowIndex = -1; // For cases with "As per Schedule" in a separate row
          let columnIndices: Record<string, string> = {};
          let jsonData: ExcelRow[] = [];
          
          // First, check if "Sub-Order-Annx-1" sheet exists, as it's likely to have our data
          const targetSheetName = workbook.SheetNames.find(
            name => name.toLowerCase().includes('sub-order') || 
                    name.toLowerCase().includes('annx') || 
                    name.toLowerCase().includes('annex')
          );
          
          // Prioritize checking the target sheet if found
          const sheetsToCheck = targetSheetName 
            ? [targetSheetName, ...workbook.SheetNames.filter(name => name !== targetSheetName)]
            : workbook.SheetNames;
          
          // Check each sheet in the workbook
          sheetLoop: for (const sheetName of sheetsToCheck) {
            const worksheet = workbook.Sheets[sheetName];
            
            // Convert to JSON
            const sheetData = XLSX.utils.sheet_to_json<ExcelRow>(worksheet, { header: 'A' });
            
            if (!Array.isArray(sheetData) || sheetData.length < 2) {
              continue; // Skip empty sheets
            }
            
            // First try: Look for headers with "As per Schedule" in the same row
            let tempHeaderRowIndex = -1;
            let tempColumnIndices: Record<string, string> = {};
            
            // Look for header row containing expected column names
            for (let i = 0; i < Math.min(50, sheetData.length); i++) {
              const row = sheetData[i];
              let foundColumns = 0;
              
              // Check each cell in this row to see if it contains our target headers
              for (const [col, value] of Object.entries(row)) {
                if (!value) continue;
                
                const valueStr = String(value).trim();
                
                // PO Sl. No. variations
                if (valueStr === 'PO Sl.No' || 
                    valueStr === 'PO Sl. No.' || 
                    valueStr === 'PO SI.No' ||
                    valueStr === 'PO Sl.No.' || 
                    valueStr === 'Sl.No') {
                  tempColumnIndices.poSlNo = col;
                  foundColumns++;
                } 
                // Item Description variations
                else if (valueStr === 'Item Description' || 
                        valueStr === 'Description' || 
                        valueStr === 'Item') {
                  tempColumnIndices.itemDescription = col;
                  foundColumns++;
                } 
                // Engineering Approval Date variations
                else if ((valueStr.includes('Engineering') && valueStr.includes('Approval') && valueStr.includes('Date')) || 
                        (valueStr.includes('Engineering') && valueStr.includes('As per Schedule'))) {
                  tempColumnIndices.engineeringApprovalDate = col;
                  foundColumns++;
                } 
                // Procurement Start Date variations
                else if ((valueStr.includes('Procurement') && valueStr.includes('Start') && valueStr.includes('Date')) || 
                        (valueStr.includes('Procurement') && valueStr.includes('As per Schedule'))) {
                  tempColumnIndices.procurementStartDate = col;
                  foundColumns++;
                } 
                // Delivery at Site variations
                else if ((valueStr.includes('Delivery') && valueStr.includes('Site')) || 
                        (valueStr.includes('Delivery') && valueStr.includes('As per Schedule'))) {
                  tempColumnIndices.deliveryAtSite = col;
                  foundColumns++;
                }
              }
              
              // If we found at least 3 of our target columns, consider this the header row
              if (foundColumns >= 3) {
                tempHeaderRowIndex = i;
                validSheetName = sheetName;
                headerRowIndex = tempHeaderRowIndex;
                columnIndices = tempColumnIndices;
                jsonData = sheetData;
                break sheetLoop;
              }
            }
            
            // Second try: Look for multi-row headers (like in the screenshot)
            // where "As per Schedule" might be in a different row
            for (let i = 0; i < Math.min(50, sheetData.length - 1); i++) {
              const row = sheetData[i];
              const nextRow = sheetData[i + 1];
              
              // First find column titles (Engineering Approval Date, etc.)
              let mainColumns: Record<string, string> = {};
              let foundMainColumns = 0;
              
              for (const [col, value] of Object.entries(row)) {
                if (!value) continue;
                
                const valueStr = String(value).trim();
                
                if (valueStr === 'PO Sl.No' || 
                    valueStr === 'PO Sl. No.' || 
                    valueStr === 'PO SI.No' ||
                    valueStr === 'PO Sl.No.' || 
                    valueStr === 'Sl.No') {
                  mainColumns.poSlNo = col;
                  foundMainColumns++;
                } 
                else if (valueStr === 'Item Description' || 
                         valueStr === 'Description' || 
                         valueStr === 'Item') {
                  mainColumns.itemDescription = col;
                  foundMainColumns++;
                }
                else if (valueStr.includes('Engineering Approval Date')) {
                  mainColumns.engineeringApprovalDate = col;
                  foundMainColumns++;
                }
                else if (valueStr.includes('Procurement Start Date')) {
                  mainColumns.procurementStartDate = col;
                  foundMainColumns++;
                }
                else if (valueStr.includes('Delivery at Site')) {
                  mainColumns.deliveryAtSite = col;
                  foundMainColumns++;
                }
              }
              
              // If we found column titles, check the next row for "As per Schedule"
              if (foundMainColumns >= 3) {
                let foundScheduleRow = false;
                
                for (const [col, value] of Object.entries(nextRow)) {
                  if (!value) continue;
                  
                  const valueStr = String(value).trim();
                  
                  if (valueStr.includes('As per Schedule')) {
                    foundScheduleRow = true;
                    
                    // Find which column this "As per Schedule" belongs to
                    if (mainColumns.engineeringApprovalDate && 
                        parseInt(col.slice(1)) === parseInt(mainColumns.engineeringApprovalDate.slice(1))) {
                      tempColumnIndices.engineeringApprovalDate = col;
                    }
                    else if (mainColumns.procurementStartDate && 
                             parseInt(col.slice(1)) === parseInt(mainColumns.procurementStartDate.slice(1))) {
                      tempColumnIndices.procurementStartDate = col;
                    }
                    else if (mainColumns.deliveryAtSite && 
                             parseInt(col.slice(1)) === parseInt(mainColumns.deliveryAtSite.slice(1))) {
                      tempColumnIndices.deliveryAtSite = col;
                    }
                  }
                }
                
                if (foundScheduleRow) {
                  // Add the simpler columns
                  if (mainColumns.poSlNo) tempColumnIndices.poSlNo = mainColumns.poSlNo;
                  if (mainColumns.itemDescription) tempColumnIndices.itemDescription = mainColumns.itemDescription;
                  
                  // If we found at least 3 total columns
                  if (Object.keys(tempColumnIndices).length >= 3) {
                    validSheetName = sheetName;
                    headerRowIndex = i;
                    subHeaderRowIndex = i + 1;
                    columnIndices = tempColumnIndices;
                    jsonData = sheetData;
                    break sheetLoop;
                  }
                }
              }
            }
          }
          
          if (!validSheetName || headerRowIndex === -1 || Object.keys(columnIndices).length < 3) {
            return {
              fileName: file.name,
              success: false,
              error: 'Required columns not found in any sheet of the Excel file',
            };
          }
          
          // Extract data from all rows after the header
          const extractedData: ExtractedData[] = [];
          const startRow = subHeaderRowIndex > -1 ? subHeaderRowIndex + 1 : headerRowIndex + 1;
          
          for (let i = startRow; i < jsonData.length; i++) {
            const row = jsonData[i];
            
            // Skip empty rows
            if (!row || Object.keys(row).length === 0) continue;
            
            // Check if this is a data row (should have PO Sl.No or Item Description)
            const hasPoSlNo = columnIndices.poSlNo && row[columnIndices.poSlNo];
            const hasItemDesc = columnIndices.itemDescription && row[columnIndices.itemDescription];
            
            if (!hasPoSlNo && !hasItemDesc) continue;
            
            // Convert all values to strings or null
            const data: ExtractedData = {
              poSlNo: columnIndices.poSlNo && row[columnIndices.poSlNo] ? String(row[columnIndices.poSlNo]) : null,
              itemDescription: columnIndices.itemDescription && row[columnIndices.itemDescription] ? String(row[columnIndices.itemDescription]) : null,
              engineeringApprovalDate: columnIndices.engineeringApprovalDate && row[columnIndices.engineeringApprovalDate] ? String(row[columnIndices.engineeringApprovalDate]) : null,
              procurementStartDate: columnIndices.procurementStartDate && row[columnIndices.procurementStartDate] ? String(row[columnIndices.procurementStartDate]) : null,
              deliveryAtSite: columnIndices.deliveryAtSite && row[columnIndices.deliveryAtSite] ? String(row[columnIndices.deliveryAtSite]) : null,
            };
            
            extractedData.push(data);
          }
          
          return {
            fileName: file.name,
            sheetName: validSheetName,
            success: true,
            data: extractedData,
          };
          
        } catch (error) {
          console.error(`Error processing file ${file.name}:`, error);
          return {
            fileName: file.name,
            success: false,
            error: 'Failed to process Excel file',
          };
        }
      })
    );
    
    // Return the results
    return NextResponse.json({
      message: 'Files processed successfully',
      files: fileResults,
    });
    
  } catch (error) {
    console.error('Error processing files:', error);
    return NextResponse.json(
      { error: 'Failed to process files' },
      { status: 500 }
    );
  }
} 