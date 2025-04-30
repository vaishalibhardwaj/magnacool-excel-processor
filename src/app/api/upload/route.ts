import { NextRequest, NextResponse } from 'next/server';

// Set the maximum file size (10MB)
export const config = {
  api: {
    bodyParser: false,
    responseLimit: '10mb',
  },
};

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
    const fileResults = await Promise.all(
      files.map(async (file) => {
        // Validate file type
        const isValidType = 
          file.type === 'application/vnd.ms-excel' || 
          file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
          
        if (!isValidType) {
          return {
            name: file.name,
            success: false,
            error: 'Invalid file type. Only Excel files (.xls, .xlsx) are allowed',
          };
        }
        
        // In a real application, you would process the Excel file here
        // For example, read its contents, validate, save to storage, etc.
        
        // Simulate processing time
        await new Promise(resolve => setTimeout(resolve, 500));
        
        return {
          name: file.name,
          size: file.size,
          success: true,
        };
      })
    );
    
    // Return the results
    return NextResponse.json({
      message: 'Files processed successfully',
      files: fileResults,
    });
    
  } catch (error) {
    console.error('Error uploading files:', error);
    return NextResponse.json(
      { error: 'Failed to process files' },
      { status: 500 }
    );
  }
} 