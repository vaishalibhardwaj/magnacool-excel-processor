'use client';

import { useState, useCallback } from 'react';
import { useDropzone } from 'react-dropzone';

export interface UploadedFile {
  file: File;
  id: string;
  status?: 'pending' | 'uploading' | 'success' | 'error';
  error?: string;
  parsedData?: ParsedData[];
  sheetName?: string;
}

interface ParsedData {
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
  data?: ParsedData[];
  sheetName?: string;
}

interface ApiResponse {
  message: string;
  files: FileResult[];
}

// Function to convert Excel date serial number to JS Date
const excelDateToJSDate = (excelDate: string | null): Date | null => {
  if (!excelDate) return null;
  
  // Try to parse the value as a number
  const serialDate = parseFloat(excelDate);
  if (isNaN(serialDate)) return null;
  
  // Excel date serial numbers start from January 1, 1900
  // And there's a leap year bug in Excel where it thinks 1900 was a leap year
  return new Date((serialDate - 25569) * 86400 * 1000);
};

// Calculate delay in days from target date to current date
const calculateDelay = (dateString: string | null): number | null => {
  if (!dateString) return null;
  
  const date = excelDateToJSDate(dateString);
  if (!date) return null;
  
  const today = new Date();
  today.setHours(0, 0, 0, 0); // Set to beginning of day for fair comparison
  
  // Return days difference (positive means past due, negative means days remaining)
  return Math.floor((today.getTime() - date.getTime()) / (1000 * 3600 * 24));
};

// Get color based on delay
const getDelayColor = (delay: number | null): string => {
  if (delay === null) return '';
  if (delay > 0) return 'bg-red-100 text-red-800'; // Past due
  if (delay === 0 || delay === -1) return 'bg-yellow-100 text-yellow-800'; // Due today or tomorrow
  return 'bg-green-100 text-green-800'; // Future date
};

// Format date for display
const formatDate = (dateString: string | null): string => {
  if (!dateString) return '-';
  
  const date = excelDateToJSDate(dateString);
  if (!date) return dateString;
  
  return date.toLocaleDateString('en-US', {
    year: 'numeric',
    month: 'short',
    day: 'numeric'
  });
};

export default function FileUpload() {
  const [files, setFiles] = useState<UploadedFile[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [isUploading, setIsUploading] = useState(false);
  const [uploadProgress, setUploadProgress] = useState(0);
  const [parsedResults, setParsedResults] = useState<ApiResponse | null>(null);

  const onDrop = useCallback((acceptedFiles: File[]) => {
    setError(null);
    
    // Check if adding these files would exceed the limit
    if (files.length + acceptedFiles.length > 20) {
      setError('Maximum 20 files allowed');
      return;
    }
    
    // Convert files to our format with IDs
    const newFiles = acceptedFiles.map(file => ({
      file,
      id: crypto.randomUUID(),
      status: 'pending' as const
    }));
    
    setFiles(prev => [...prev, ...newFiles]);
  }, [files]);

  const removeFile = (id: string) => {
    setFiles(files.filter(file => file.id !== id));
    // Clear parsed results when files change
    setParsedResults(null);
  };

  const handleUpload = async () => {
    if (files.length === 0) {
      setError('Please select at least one file');
      return;
    }

    if (isUploading) return;

    setIsUploading(true);
    setError(null);
    setParsedResults(null);
    
    try {
      // Update all files to uploading status
      setFiles(prev => 
        prev.map(file => ({
          ...file,
          status: 'uploading'
        }))
      );

      // Create form data
      const formData = new FormData();
      files.forEach(fileObj => {
        formData.append('files', fileObj.file);
      });

      // Simulate upload progress
      const progressInterval = setInterval(() => {
        setUploadProgress(prev => {
          const newProgress = prev + Math.random() * 15;
          return newProgress > 95 ? 95 : newProgress;
        });
      }, 500);

      // Send to the Excel parsing API
      const response = await fetch('/api/parse-excel', {
        method: 'POST',
        body: formData,
      });

      clearInterval(progressInterval);
      setUploadProgress(100);

      if (!response.ok) {
        const errorData = await response.json() as { error: string };
        throw new Error(errorData.error || 'Upload failed');
      }

      const result = await response.json() as ApiResponse;
      
      // Update file statuses based on results
      setFiles(prev => 
        prev.map(fileObj => {
          const fileResult = result.files.find((r) => r.fileName === fileObj.file.name);
          return {
            ...fileObj,
            status: fileResult?.success ? 'success' : 'error',
            error: fileResult?.error,
            parsedData: fileResult?.data,
            sheetName: fileResult?.sheetName
          };
        })
      );

      // Show the parsed results
      setParsedResults(result);
      
      // Show success message
      console.log('Upload successful:', result);
    } catch (err: unknown) {
      const errorMessage = err instanceof Error ? err.message : 'Failed to upload files';
      setError(errorMessage);
      
      // Mark all as error
      setFiles(prev => 
        prev.map(file => ({
          ...file,
          status: 'error'
        }))
      );
    } finally {
      setIsUploading(false);
      // Reset progress after a delay
      setTimeout(() => setUploadProgress(0), 2000);
    }
  };

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      'application/vnd.ms-excel': ['.xls'],
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx']
    },
    multiple: true,
    disabled: isUploading
  });

  return (
    <div className="w-full max-w-4xl mx-auto">
      <div 
        {...getRootProps()} 
        className={`p-10 border-2 border-dashed rounded-lg text-center cursor-pointer transition-colors ${
          isDragActive 
            ? 'border-blue-500 bg-blue-50' 
            : isUploading
              ? 'border-gray-300 bg-gray-100 cursor-not-allowed'
              : 'border-gray-300 hover:border-gray-400'
        }`}
      >
        <input {...getInputProps()} />
        <div className="space-y-4">
          <div className="flex justify-center">
            <svg xmlns="http://www.w3.org/2000/svg" className="h-12 w-12 text-gray-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
            </svg>
          </div>
          <div>
            <p className="text-lg font-medium text-gray-900">
              {isDragActive 
                ? 'Drop Excel files here' 
                : isUploading 
                  ? 'Uploading files...' 
                  : 'Drag & drop Excel files here'}
            </p>
            <p className="text-sm text-gray-500 mt-1">
              {!isUploading && 'or click to browse files'}
            </p>
            <p className="text-xs text-gray-400 mt-2">Only .xls and .xlsx files accepted (max 20 files)</p>
          </div>
        </div>
      </div>

      {error && (
        <div className="mt-4 p-3 bg-red-50 text-red-600 rounded-md">
          {error}
        </div>
      )}
      
      {isUploading && (
        <div className="mt-4">
          <div className="w-full bg-gray-200 rounded-full h-2.5">
            <div 
              className="bg-blue-600 h-2.5 rounded-full transition-all duration-300" 
              style={{ width: `${uploadProgress}%` }}
            ></div>
          </div>
          <p className="text-sm text-gray-500 mt-2 text-center">
            Uploading... {Math.round(uploadProgress)}%
          </p>
        </div>
      )}

      {files.length > 0 && (
        <div className="mt-6">
          <h3 className="text-lg font-medium text-gray-900 mb-3">Selected Files ({files.length}/20)</h3>
          <ul className="space-y-2">
            {files.map((file) => (
              <li key={file.id} className="flex items-center justify-between p-3 bg-white border rounded-md">
                <div className="flex items-center flex-1 mr-4">
                  <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5 text-green-600 mr-2" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 17v-2m3 2v-4m3 4v-6m2 10H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                  </svg>
                  <div className="flex flex-col">
                    <span className="text-sm text-gray-700 truncate max-w-xs">{file.file.name}</span>
                    <span className="text-xs text-gray-500">
                      {(file.file.size / 1024 / 1024).toFixed(2)} MB
                    </span>
                    {file.error && (
                      <span className="text-xs text-red-500">{file.error}</span>
                    )}
                    {file.sheetName && (
                      <span className="text-xs text-green-600">Sheet: {file.sheetName}</span>
                    )}
                  </div>
                </div>
                <div className="flex items-center">
                  {file.status === 'uploading' && (
                    <div className="mr-2">
                      <svg className="animate-spin h-4 w-4 text-blue-500" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                        <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                        <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
                      </svg>
                    </div>
                  )}
                  {file.status === 'success' && (
                    <div className="mr-2 text-green-500">
                      <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
                      </svg>
                    </div>
                  )}
                  {file.status === 'error' && (
                    <div className="mr-2 text-red-500">
                      <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
                      </svg>
                    </div>
                  )}
                  {!isUploading && (
                    <button
                      onClick={(e) => {
                        e.stopPropagation();
                        removeFile(file.id);
                      }}
                      className="text-red-500 hover:text-red-700 focus:outline-none"
                      disabled={isUploading}
                    >
                      <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" />
                      </svg>
                    </button>
                  )}
                </div>
              </li>
            ))}
          </ul>
          
          <div className="mt-6">
            <button
              onClick={handleUpload}
              disabled={isUploading || files.length === 0}
              className={`w-full py-2 px-4 rounded-md font-medium text-white transition-colors ${
                isUploading || files.length === 0
                  ? 'bg-blue-300 cursor-not-allowed'
                  : 'bg-blue-600 hover:bg-blue-700'
              }`}
            >
              {isUploading ? 'Processing...' : 'Process Excel Files'}
            </button>
          </div>
        </div>
      )}
      
      {/* Display Parsed Results with Delay Calculations */}
      {parsedResults && (
        <div className="mt-8 border border-gray-200 rounded-lg overflow-hidden">
          <div className="p-4 bg-gray-50 border-b border-gray-200">
            <h2 className="text-xl font-semibold text-gray-800">Parsed Excel Data with Delay Analysis</h2>
            <div className="mt-2 text-sm text-gray-600">
              <div className="flex items-center space-x-6">
                <div className="flex items-center">
                  <div className="w-4 h-4 bg-green-100 mr-2"></div>
                  <span>On Schedule (2+ days remaining)</span>
                </div>
                <div className="flex items-center">
                  <div className="w-4 h-4 bg-yellow-100 mr-2"></div>
                  <span>Due Soon (Today or Tomorrow)</span>
                </div>
                <div className="flex items-center">
                  <div className="w-4 h-4 bg-red-100 mr-2"></div>
                  <span>Overdue</span>
                </div>
              </div>
            </div>
          </div>
          
          <div className="overflow-x-auto">
            {parsedResults.files.map((fileResult, fileIndex) => (
              fileResult.success && fileResult.data && fileResult.data.length > 0 ? (
                <div key={fileIndex} className="p-4 border-b border-gray-200">
                  <h3 className="text-lg font-medium text-gray-700 mb-2">
                    {fileResult.fileName}
                    {fileResult.sheetName && (
                      <span className="ml-2 text-sm text-green-600">
                        (Sheet: {fileResult.sheetName})
                      </span>
                    )}
                  </h3>
                  <div className="overflow-x-auto">
                    <table className="min-w-full divide-y divide-gray-200">
                      <thead className="bg-gray-50">
                        <tr>
                          <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">PO Sl. No.</th>
                          <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Item Description</th>
                          <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Engineering Approval Date</th>
                          <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Eng. Approval Delay</th>
                          <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Procurement Start Date</th>
                          <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Procurement Delay</th>
                          <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Delivery at Site</th>
                          <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Delivery Delay</th>
                        </tr>
                      </thead>
                      <tbody className="bg-white divide-y divide-gray-200">
                        {fileResult.data.map((item, index) => {
                          // Calculate delays
                          const engDelay = calculateDelay(item.engineeringApprovalDate);
                          const procDelay = calculateDelay(item.procurementStartDate);
                          const deliveryDelay = calculateDelay(item.deliveryAtSite);
                          
                          // Get color classes
                          const engDelayColor = getDelayColor(engDelay);
                          const procDelayColor = getDelayColor(procDelay);
                          const deliveryDelayColor = getDelayColor(deliveryDelay);
                          
                          return (
                            <tr key={index} className={index % 2 === 0 ? 'bg-white' : 'bg-gray-50'}>
                              <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">{item.poSlNo || '-'}</td>
                              <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900 font-medium">{item.itemDescription || '-'}</td>
                              
                              {/* Engineering Approval Date */}
                              <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                                {formatDate(item.engineeringApprovalDate)}
                              </td>
                              <td className={`px-6 py-4 whitespace-nowrap text-sm ${engDelayColor} rounded`}>
                                {engDelay !== null ? `${engDelay} days` : '-'}
                              </td>
                              
                              {/* Procurement Start Date */}
                              <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                                {formatDate(item.procurementStartDate)}
                              </td>
                              <td className={`px-6 py-4 whitespace-nowrap text-sm ${procDelayColor} rounded`}>
                                {procDelay !== null ? `${procDelay} days` : '-'}
                              </td>
                              
                              {/* Delivery at Site */}
                              <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                                {formatDate(item.deliveryAtSite)}
                              </td>
                              <td className={`px-6 py-4 whitespace-nowrap text-sm ${deliveryDelayColor} rounded`}>
                                {deliveryDelay !== null ? `${deliveryDelay} days` : '-'}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
              ) : fileResult.success && (!fileResult.data || fileResult.data.length === 0) ? (
                <div key={fileIndex} className="p-4 border-b border-gray-200">
                  <h3 className="text-lg font-medium text-gray-700 mb-2">{fileResult.fileName}</h3>
                  <p className="text-gray-500">No data was extracted from this file.</p>
                </div>
              ) : null
            ))}
          </div>
        </div>
      )}
    </div>
  );
} 