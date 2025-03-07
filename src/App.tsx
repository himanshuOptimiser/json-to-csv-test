import React from 'react';
import * as XLSX from 'xlsx';
import { Parser } from '@json2csv/plainjs';
import { Customer } from './types';
import { 
  createTheme,
  ThemeProvider,
  Container,
  Box,
} from '@mui/material';
import FileDownloadIcon from '@mui/icons-material/FileDownload';
import { BrowserRouter, Routes, Route, Link } from 'react-router-dom';
import { CustomDataExport } from './components/CustomDataExport';

const theme = createTheme({
  typography: {
    allVariants: {
      color: '#1a1a1a',
    },
  },
  components: {
    MuiTableCell: {
      styleOverrides: {
        root: {
          color: '#1a1a1a',
        },
      },
    },
    MuiContainer: {
      styleOverrides: {
        root: {
          backgroundColor: '#ffffff',
          minHeight: '100vh',
        },
      },
    },
  },
});

const App = () => {
  // Sample nested data
  const data: Customer[] = [
    {
      "events": [
        {
          "eventId": "EVT20250305-001",
          "eventName": "TechConnect 2025",
          "eventType": "Conference",
          "startDate": "2025-09-15T09:00:00Z",
          "endDate": "2025-09-17T18:00:00Z",
          "venue": {
            "name": "Global Convention Center",
            "address": {
              "street": "123 Innovation Avenue",
              "city": "San Francisco",
              "state": "CA",
              "zipCode": "94105",
              "country": "USA"
            },
            "capacity": 5000
          },
          "organizer": {
            "name": "TechEvents Inc.",
            "contactPerson": {
              "firstName": "Jane",
              "lastName": "Smith",
              "email": "jane.smith@techevents.com",
              "phone": "+1-555-123-4567"
            }
          },
          "ticketTypes": [
            {
              "type": "Early Bird",
              "price": 299.99,
              "availableUntil": "2025-06-30T23:59:59Z"
            },
            {
              "type": "Regular",
              "price": 399.99,
              "availableUntil": "2025-09-14T23:59:59Z"
            },
            {
              "type": "VIP",
              "price": 699.99,
              "limit": 100
            }
          ],
          "sessions": [
            {
              "sessionId": "S001",
              "title": "Future of AI in Business",
              "speaker": "Dr. Alan Turing",
              "startTime": "2025-09-15T10:00:00Z",
              "endTime": "2025-09-15T11:30:00Z",
              "room": "Ballroom A"
            },
            {
              "sessionId": "S002",
              "title": "Blockchain Revolution",
              "speaker": "Satoshi Nakamoto",
              "startTime": "2025-09-16T14:00:00Z",
              "endTime": "2025-09-16T15:30:00Z",
              "room": "Ballroom B"
            }
          ],
          "attendees": [
            {
              "attendeeId": "ATT001",
              "firstName": "John",
              "lastName": "Doe",
              "email": "john.doe@example.com",
              "ticketType": "Early Bird",
              "registrationDate": "2025-05-15T08:30:00Z",
              "preferences": {
                "dietaryRestrictions": ["Vegetarian"],
                "sessionInterests": ["S001", "S002"],
                "networkingInterests": ["AI", "Blockchain"]
              },
              "payment": {
                "amount": 299.99,
                "method": "Credit Card",
                "transactionId": "TXN123456"
              }
            },
            {
              "attendeeId": "ATT002",
              "firstName": "Alice",
              "lastName": "Johnson",
              "email": "alice.johnson@example.com",
              "ticketType": "VIP",
              "registrationDate": "2025-07-20T14:45:00Z",
              "preferences": {
                "dietaryRestrictions": ["Gluten-free"],
                "sessionInterests": ["S002"],
                "networkingInterests": ["Blockchain", "Cybersecurity"]
              },
              "payment": {
                "amount": 699.99,
                "method": "PayPal",
                "transactionId": "TXN789012"
              }
            }
          ],
          "sponsors": [
            {
              "name": "TechGiant Corp",
              "level": "Platinum",
              "boothNumber": "B001",
              "logo": "https://example.com/techgiant-logo.png"
            },
            {
              "name": "InnovateSoft",
              "level": "Gold",
              "boothNumber": "B010",
              "logo": "https://example.com/innovatesoft-logo.png"
            }
          ]
        },
        {
          "eventId": "EVT20250410-002",
          "eventName": "Green Energy Symposium",
          "eventType": "Symposium",
          "startDate": "2025-10-05T08:00:00Z",
          "endDate": "2025-10-07T17:00:00Z",
          "venue": {
            "name": "EcoTech Center",
            "address": {
              "street": "789 Sustainability Road",
              "city": "Portland",
              "state": "OR",
              "zipCode": "97201",
              "country": "USA"
            },
            "capacity": 3000
          },
          "organizer": {
            "name": "GreenFuture Organization",
            "contactPerson": {
              "firstName": "Michael",
              "lastName": "Green",
              "email": "michael.green@greenfuture.org",
              "phone": "+1-555-987-6543"
            }
          },
          "ticketTypes": [
            {
              "type": "Student",
              "price": 149.99,
              "availableUntil": "2025-09-30T23:59:59Z"
            },
            {
              "type": "Professional",
              "price": 299.99,
              "availableUntil": "2025-10-04T23:59:59Z"
            }
          ],
          "sessions": [
            {
              "sessionId": "GES001",
              "title": "Solar Energy Innovations",
              "speaker": "Dr. Sun Bright",
              "startTime": "2025-10-05T09:00:00Z",
              "endTime": "2025-10-05T10:30:00Z",
              "room": "Solar Hall"
            },
            {
              "sessionId": "GES002",
              "title": "Wind Power Technologies",
              "speaker": "Prof. Gale Force",
              "startTime": "2025-10-06T11:00:00Z",
              "endTime": "2025-10-06T12:30:00Z",
              "room": "Breeze Auditorium"
            }
          ],
          "attendees": [
            {
              "attendeeId": "GATT001",
              "firstName": "Emma",
              "lastName": "Watson",
              "email": "emma.watson@example.com",
              "ticketType": "Professional",
              "registrationDate": "2025-08-01T10:15:00Z",
              "preferences": {
                "dietaryRestrictions": ["Vegan"],
                "sessionInterests": ["GES001", "GES002"],
                "networkingInterests": ["Solar", "Wind"]
              },
              "payment": {
                "amount": 299.99,
                "method": "Bank Transfer",
                "transactionId": "GTXN567890"
              }
            },
            {
              "attendeeId": "GATT002",
              "firstName": "Oliver",
              "lastName": "Green",
              "email": "oliver.green@example.com",
              "ticketType": "Student",
              "registrationDate": "2025-09-15T09:30:00Z",
              "preferences": {
                "dietaryRestrictions": ["Nut Allergy"],
                "sessionInterests": ["GES001"],
                "networkingInterests": ["Solar", "Energy Storage"]
              },
              "payment": {
                "amount": 149.99,
                "method": "Credit Card",
                "transactionId": "GTXN098765"
              }
            }
          ],
          "sponsors": [
            {
              "name": "SolarPower Co.",
              "level": "Gold",
              "boothNumber": "GB001",
              "logo": "https://example.com/solarpower-logo.png"
            },
            {
              "name": "WindTech Industries",
              "level": "Silver",
              "boothNumber": "GB005",
              "logo": "https://example.com/windtech-logo.png"
            }
          ]
        }
      ]
    }
  ];

  const prepareDataForExport = (data: any[]) => {
    const attendeesData: any[] = [];
    const sessionsData: any[] = [];
    const sponsorsData: any[] = [];
    const ticketTypesData: any[] = [];
    const eventsData: any[] = [];

    data[0].events.forEach((event: any) => {
      // Base event information
      eventsData.push({
        'Event ID': event.eventId,
        'Event Name': event.eventName,
        'Event Type': event.eventType,
        'Start Date': event.startDate,
        'End Date': event.endDate,
        'Venue Name': event.venue.name,
        'Venue Address': `${event.venue.address.street}, ${event.venue.address.city}`,
        'Venue State': event.venue.address.state,
        'Venue Country': event.venue.address.country,
        'Venue Capacity': event.venue.capacity,
        'Organizer': event.organizer.name,
        'Organizer Contact': `${event.organizer.contactPerson.firstName} ${event.organizer.contactPerson.lastName}`,
        'Organizer Email': event.organizer.contactPerson.email,
        'Organizer Phone': event.organizer.contactPerson.phone
      });

      // Flatten attendees
      event.attendees.forEach((attendee: any) => {
        attendeesData.push({
          'Event ID': event.eventId,
          'Event Name': event.eventName,
          'Attendee ID': attendee.attendeeId,
          'First Name': attendee.firstName,
          'Last Name': attendee.lastName,
          'Email': attendee.email,
          'Ticket Type': attendee.ticketType,
          'Registration Date': attendee.registrationDate,
          'Dietary Restrictions': attendee.preferences.dietaryRestrictions.join(', '),
          'Session Interests': attendee.preferences.sessionInterests.join(', '),
          'Networking Interests': attendee.preferences.networkingInterests.join(', '),
          'Payment Amount': attendee.payment.amount,
          'Payment Method': attendee.payment.method,
          'Transaction ID': attendee.payment.transactionId
        });
      });

      // Flatten sessions
      event.sessions.forEach((session: any) => {
        sessionsData.push({
          'Event ID': event.eventId,
          'Event Name': event.eventName,
          'Session ID': session.sessionId,
          'Title': session.title,
          'Speaker': session.speaker,
          'Start Time': session.startTime,
          'End Time': session.endTime,
          'Room': session.room
        });
      });

      // Flatten sponsors
      event.sponsors.forEach((sponsor: any) => {
        sponsorsData.push({
          'Event ID': event.eventId,
          'Event Name': event.eventName,
          'Sponsor Name': sponsor.name,
          'Level': sponsor.level,
          'Booth Number': sponsor.boothNumber,
          'Logo URL': sponsor.logo
        });
      });

      // Flatten ticket types
      event.ticketTypes.forEach((ticket: any) => {
        ticketTypesData.push({
          'Event ID': event.eventId,
          'Event Name': event.eventName,
          'Ticket Type': ticket.type,
          'Price': ticket.price,
          'Available Until': ticket.availableUntil,
          'Limit': ticket.limit || 'Unlimited'
        });
      });
    });

    return {
      events: eventsData,
      attendees: attendeesData,
      sessions: sessionsData,
      sponsors: sponsorsData,
      ticketTypes: ticketTypesData
    };
  };

  const exportToExcel = () => {
    try {
      const preparedData = prepareDataForExport(data);
      const wb = XLSX.utils.book_new();

      // Helper function to auto-size columns
      const autoFitColumns = (json: any[], worksheet: XLSX.WorkSheet) => {
        const objectMaxLength: { [key: string]: number } = {};
        
        // Get maximum length of each column
        json.forEach(obj => {
          Object.entries(obj).forEach(([key, value]) => {
            const valueLength = String(value).length;
            objectMaxLength[key] = Math.max(
              valueLength,
              objectMaxLength[key] || key.length
            );
          });
        });

        // Set column widths
        const cols: XLSX.ColInfo[] = [];
        Object.entries(objectMaxLength).forEach(([key, length]) => {
          cols.push({
            wch: Math.min(length + 2, 50) // Add padding, max width 50
          });
        });
        worksheet['!cols'] = cols;
      };

      // Create and style each worksheet
      Object.entries(preparedData).forEach(([sheetName, sheetData]) => {
        const ws = XLSX.utils.json_to_sheet(sheetData);

        // Style the header row
        const headerRange = XLSX.utils.decode_range(ws['!ref'] || 'A1');
        for (let col = headerRange.s.c; col <= headerRange.e.c; col++) {
          const cellRef = XLSX.utils.encode_cell({ r: 0, c: col });
          if (!ws[cellRef]) continue;

          ws[cellRef].s = {
            font: {
              bold: true,
              color: { rgb: "FFFFFF" }
            },
            fill: {
              fgColor: { rgb: "4472C4" }
            },
            alignment: {
              horizontal: 'center',
              vertical: 'center'
            }
          };
        }

        // Auto-fit columns
        autoFitColumns(sheetData, ws);

        // Add borders to all cells
        const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
        for (let row = range.s.r; row <= range.e.r; row++) {
          for (let col = range.s.c; col <= range.e.c; col++) {
            const cellRef = XLSX.utils.encode_cell({ r: row, c: col });
            if (!ws[cellRef]) {
              ws[cellRef] = { v: '' };
            }
            ws[cellRef].s = {
              ...ws[cellRef].s,
              border: {
                top: { style: 'thin' },
                bottom: { style: 'thin' },
                left: { style: 'thin' },
                right: { style: 'thin' }
              }
            };
          }
        }

        // Add the worksheet to workbook
        XLSX.utils.book_append_sheet(
          wb, 
          ws, 
          sheetName.charAt(0).toUpperCase() + sheetName.slice(1)
        );
      });

      // Set file properties
      wb.Props = {
        Title: "Event Complete Report",
        Subject: "Event Data Export",
        Author: "Event Management System",
        CreatedDate: new Date()
      };

      // Generate Excel file
      XLSX.writeFile(wb, 'event-complete-report.xlsx');
    } catch (err) {
      console.error('Error exporting Excel:', err);
    }
  };

  return (
    <ThemeProvider theme={theme}>
      <Box sx={{ backgroundColor: '#f5f5f5', minHeight: '100vh', width: '100%' }}>
        <BrowserRouter>
          <Container maxWidth={false} sx={{ 
            minHeight: '100vh',
            backgroundColor: '#ffffff',
            p: { xs: 1, sm: 2, md: 3 }, // Responsive padding
            boxShadow: 1 
          }}>
            <Box sx={{ 
              display: 'flex', 
              gap: 2, 
              mb: 3 
            }}>
              <Link to="/" style={{ 
                textDecoration: 'none',
                padding: '8px 16px',
                backgroundColor: '#1976d2',
                color: 'white',
                borderRadius: '4px'
              }}>
                Sample Data Export
              </Link>
              <Link to="/custom" style={{ 
                textDecoration: 'none',
                padding: '8px 16px',
                backgroundColor: '#1976d2',
                color: 'white',
                borderRadius: '4px'
              }}>
                Custom Data Export
              </Link>
            </Box>
            <Routes>
              <Route path="/" element={
                <Box sx={{ p: 3 }}>
                  <h1 style={{ color: '#1a1a1a', marginBottom: '20px' }}>Event Data Export</h1>
                  <button 
                    onClick={exportToExcel}
                    style={{
                      backgroundColor: '#1976d2',
                      color: 'white',
                      padding: '10px 20px',
                      border: 'none',
                      borderRadius: '4px',
                      cursor: 'pointer',
                      marginBottom: '20px'
                    }}
                  >
                    Export to Excel
                  </button>
                  
                  <h2 style={{ color: '#1a1a1a', marginBottom: '10px' }}>Preview Data:</h2>
                  <Box component="pre" sx={{ 
                    bgcolor: 'grey.100', 
                    p: 2, 
                    borderRadius: 1,
                    overflow: 'auto',
                    maxHeight: '400px',
                    color: '#1a1a1a',
                  }}>
                    {JSON.stringify(data, null, 2)}
                  </Box>
                </Box>
              } />
              <Route path="/custom" element={<CustomDataExport />} />
            </Routes>
          </Container>
        </BrowserRouter>
      </Box>
    </ThemeProvider>
  );
};

export default App;
