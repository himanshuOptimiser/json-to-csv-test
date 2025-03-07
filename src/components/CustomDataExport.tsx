import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { 
  Box, 
  TextField, 
  Button, 
  Alert, 
  Typography,
  CircularProgress,
  Grid,
  Paper
} from '@mui/material';
import FileDownloadIcon from '@mui/icons-material/FileDownload';

export const CustomDataExport = () => {
  const [jsonData, setJsonData] = useState('');
  const [error, setError] = useState('');
  const [isLoading, setIsLoading] = useState(false);

  const processNestedData = (data: any) => {
    // Initialize sheets object to store different types of data
    const sheets: { [key: string]: any[] } = {};

    const processArray = (arr: any[], prefix: string = '') => {
      // Create a sheet name from the prefix or use 'Main' if no prefix
      const sheetName = prefix || 'Main';
      if (!sheets[sheetName]) {
        sheets[sheetName] = [];
      }

      arr.forEach(item => {
        // Process current level object
        const flatObject = Object.entries(item).reduce((acc: any, [key, value]) => {
          if (!Array.isArray(value) && typeof value !== 'object') {
            acc[key] = value;
          } else if (Array.isArray(value)) {
            // For arrays, create a new sheet and store reference
            processArray(value, key);
            // Store count in parent
            acc[`${key}_count`] = value.length;
          } else if (value && typeof value === 'object') {
            // Flatten nested objects
            Object.entries(value).forEach(([nestedKey, nestedValue]) => {
              acc[`${key}_${nestedKey}`] = nestedValue;
            });
          }
          return acc;
        }, {});
        
        sheets[sheetName].push(flatObject);
      });
    };

    // Start processing from the root
    if (Array.isArray(data)) {
      processArray(data);
    } else {
      processArray([data]);
    }

    return sheets;
  };

  const handleExport = async () => {
    try {
      setIsLoading(true);
      setError('');
      
      const parsedData = JSON.parse(jsonData);
      const wb = XLSX.utils.book_new();

      // Process the nested data into separate sheets
      const processedSheets = processNestedData(parsedData);

      // Create worksheets for each data type
      Object.entries(processedSheets).forEach(([sheetName, sheetData]) => {
        const ws = XLSX.utils.json_to_sheet(sheetData);

        // Auto-fit columns
        const maxWidths: { [key: string]: number } = {};
        sheetData.forEach(row => {
          Object.entries(row).forEach(([key, value]) => {
            const valueLength = String(value).length;
            maxWidths[key] = Math.max(maxWidths[key] || 0, valueLength, key.length);
          });
        });

        ws['!cols'] = Object.values(maxWidths).map(width => ({
          wch: Math.min(width + 2, 50)
        }));

        // Style headers
        const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
        for (let col = range.s.c; col <= range.e.c; col++) {
          const cellRef = XLSX.utils.encode_cell({ r: 0, c: col });
          if (!ws[cellRef]) continue;

          ws[cellRef].s = {
            font: { bold: true, color: { rgb: "FFFFFF" } },
            fill: { fgColor: { rgb: "4472C4" } },
            alignment: { horizontal: 'center', vertical: 'center' }
          };
        }

        // Add the worksheet to workbook
        XLSX.utils.book_append_sheet(
          wb,
          ws,
          sheetName.charAt(0).toUpperCase() + sheetName.slice(1)
        );
      });

      // Generate Excel file
      XLSX.writeFile(wb, 'custom-data-export.xlsx');
      setError('');
    } catch (err) {
      setError('Invalid JSON format. Please check your input.');
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <Box sx={{ 
      backgroundColor: '#f8f9fa',
      borderRadius: 2,
      minHeight: '80vh',
      width: '100%',
      p: { xs: 1, sm: 2, md: 3 } // Responsive padding
    }}>
      <Grid container spacing={{ xs: 2, md: 4 }} sx={{ width: '100%', m: 0 }}>
        {/* Left side - Export Form */}
        <Grid item xs={12} md={7} sx={{ p: { xs: 1, sm: 2 } }}>
          <Paper elevation={0} sx={{ 
            p: 4, 
            backgroundColor: 'white',
            borderRadius: 2,
            height: '100%'
          }}>
            <Typography variant="h4" gutterBottom sx={{ 
              fontWeight: 600,
              color: '#2c3e50'
            }}>
              Custom Data Export
            </Typography>
            <Typography variant="body1" sx={{ 
              mb: 3,
              color: '#34495e'
            }}>
              Paste your JSON data below to export it to Excel
            </Typography>
            {error && (
              <Alert severity="error" sx={{ mb: 2 }}>
                {error}
              </Alert>
            )}
            <TextField
              multiline
              rows={12}
              fullWidth
              value={jsonData}
              onChange={(e) => setJsonData(e.target.value)}
              placeholder="Paste your JSON data here..."
              sx={{ 
                mb: 3,
                backgroundColor: '#ffffff',
                '& .MuiOutlinedInput-root': {
                  backgroundColor: '#ffffff',
                }
              }}
            />
            <Button
              variant="contained"
              onClick={handleExport}
              startIcon={isLoading ? <CircularProgress size={20} color="inherit" /> : <FileDownloadIcon />}
              disabled={!jsonData || isLoading}
              sx={{
                backgroundColor: '#3498db',
                '&:hover': {
                  backgroundColor: '#2980b9'
                }
              }}
            >
              {isLoading ? 'Processing...' : 'Export to Excel'}
            </Button>
          </Paper>
        </Grid>

        {/* Right side - Image and Quote */}
        <Grid item xs={12} md={5} sx={{ p: { xs: 1, sm: 2 } }}>
          <Paper elevation={0} sx={{ 
            p: 4, 
            backgroundColor: '#34495e',
            borderRadius: 2,
            height: '100%',
            display: 'flex',
            flexDirection: 'column',
            justifyContent: 'center',
            alignItems: 'center',
            textAlign: 'center',
            color: 'white'
          }}>
            <img 
              src="https://images.unsplash.com/photo-1548438294-1ad5d5f4f063?q=80&w=2072&auto=format&fit=crop&ixlib=rb-4.0.3&ixid=M3wxMjA3fDB8MHxwaG90by1wYWdlfHx8fGVufDB8fHx8fA%3D%3D" 
              alt="Data Processing Illustration"
              style={{
                width: '80%',
                maxWidth: '300px',
                marginBottom: '2rem'
              }}
            />
            <Typography variant="h5" sx={{ 
              fontWeight: 500,
              mb: 2,
              fontStyle: 'italic'
            }}>
              "Transform your data into insights"
            </Typography>
            <Typography variant="body1" sx={{ 
              opacity: 0.9,
              maxWidth: '80%'
            }}>
              Easily convert your JSON data into beautifully formatted Excel spreadsheets with just a few clicks.
            </Typography>
          </Paper>
        </Grid>
      </Grid>
    </Box>
  );
};
