"use client";

import { useState, useRef } from 'react';
import * as XLSX from 'xlsx';

interface ConversionStatus {
  type: 'idle' | 'loading' | 'success' | 'error';
  message?: string;
}

export default function SpreadsheetConverter() {
  const [file, setFile] = useState<File | null>(null);
  const [status, setStatus] = useState<ConversionStatus>({ type: 'idle' });
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = e.target.files?.[0];
    if (selectedFile) {
      const validTypes = [
        'application/vnd.ms-excel',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.oasis.opendocument.spreadsheet'
      ];
      
      if (validTypes.includes(selectedFile.type) || selectedFile.name.endsWith('.xlsx') || selectedFile.name.endsWith('.xls')) {
        setFile(selectedFile);
        setStatus({ type: 'idle' });
      } else {
        setStatus({ type: 'error', message: 'Please select a valid Excel file (.xlsx or .xls)' });
      }
    }
  };

  const convertToICS = async () => {
    if (!file) return;

    setStatus({ type: 'loading', message: 'Converting spreadsheet...' });

    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: 'array' });
      
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData = XLSX.utils.sheet_to_json(firstSheet);

      const icsContent = generateICS(jsonData);

      const blob = new Blob([icsContent], { type: 'text/calendar' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${file.name.replace(/\.[^/.]+$/, '')}_calendar.ics`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

      setStatus({ type: 'success', message: 'Calendar file downloaded successfully!' });
    } catch (error) {
      console.error('Conversion error:', error);
      setStatus({ type: 'error', message: 'Failed to convert file. Please check the format.' });
    }
  };

  const generateICS = (data: any[]): string => {
    let ics = 'BEGIN:VCALENDAR\r\n';
    ics += 'VERSION:2.0\r\n';
    ics += 'PRODID:-//Spreadsheet Converter//EN\r\n';
    ics += 'CALSCALE:GREGORIAN\r\n';
    ics += 'METHOD:PUBLISH\r\n';

    data.forEach((row, index) => {
      const title = row.title || row.Title || row.event || row.Event || row.subject || row.Subject || `Event ${index + 1}`;
      const description = row.description || row.Description || row.notes || row.Notes || '';
      const location = row.location || row.Location || row.place || row.Place || '';
      
      let startDate = row.start || row.Start || row.date || row.Date || row.startDate || row.StartDate;
      let endDate = row.end || row.End || row.endDate || row.EndDate;

      if (startDate) {
        const start = formatDateForICS(startDate);
        const end = endDate ? formatDateForICS(endDate) : start;

        ics += 'BEGIN:VEVENT\r\n';
        ics += `UID:${Date.now()}-${index}@spreadsheet-converter\r\n`;
        ics += `DTSTAMP:${formatDateForICS(new Date())}\r\n`;
        ics += `DTSTART:${start}\r\n`;
        ics += `DTEND:${end}\r\n`;
        ics += `SUMMARY:${escapeICSText(title)}\r\n`;
        if (description) {
          ics += `DESCRIPTION:${escapeICSText(description)}\r\n`;
        }
        if (location) {
          ics += `LOCATION:${escapeICSText(location)}\r\n`;
        }
        ics += 'END:VEVENT\r\n';
      }
    });

    ics += 'END:VCALENDAR\r\n';
    return ics;
  };

  const formatDateForICS = (date: any): string => {
    let d: Date;
    
    if (date instanceof Date) {
      d = date;
    } else if (typeof date === 'number') {
      d = new Date((date - 25569) * 86400 * 1000);
    } else {
      d = new Date(date);
    }

    if (isNaN(d.getTime())) {
      d = new Date();
    }

    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    const hours = String(d.getHours()).padStart(2, '0');
    const minutes = String(d.getMinutes()).padStart(2, '0');
    const seconds = String(d.getSeconds()).padStart(2, '0');

    return `${year}${month}${day}T${hours}${minutes}${seconds}`;
  };

  const escapeICSText = (text: string): string => {
    return String(text)
      .replace(/\\/g, '\\\\')
      .replace(/;/g, '\\;')
      .replace(/,/g, '\\,')
      .replace(/\n/g, '\\n');
  };

  const reset = () => {
    setFile(null);
    setStatus({ type: 'idle' });
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  return (
    <div className="xp-window">
      <div className="xp-titlebar">
        <div className="xp-titlebar-text">
          <svg width="16" height="16" viewBox="0 0 16 16" fill="none" style={{ filter: 'drop-shadow(1px 1px 1px rgba(0,0,0,0.3))' }}>
            <rect x="2" y="3" width="12" height="10" rx="1" fill="#ffffff"/>
            <rect x="3" y="5" width="10" height="1" fill="#0054e3"/>
            <rect x="3" y="7" width="10" height="1" fill="#0054e3"/>
            <rect x="3" y="9" width="10" height="1" fill="#0054e3"/>
            <rect x="3" y="11" width="6" height="1" fill="#0054e3"/>
          </svg>
          <span>Spreadsheet to Calendar Converter</span>
        </div>
        <div className="xp-titlebar-buttons">
          <div className="xp-button-minimize">_</div>
          <div className="xp-button-maximize">□</div>
          <div className="xp-button-close">×</div>
        </div>
      </div>

      <div className="xp-content">
        {/* Main groupbox */}
        <div className="xp-groupbox">
          <div className="xp-groupbox-title">File Selection</div>
          
          <div style={{ marginBottom: '16px' }}>
            <label className="block text-xs mb-2" style={{ color: '#000' }}>
              Select an Excel spreadsheet to convert:
            </label>
            
            <div className="flex gap-2 items-center">
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls"
                onChange={handleFileSelect}
                className="hidden"
                id="file-input"
              />
              <label htmlFor="file-input">
                <span className="xp-button inline-flex items-center gap-2">
                  <svg width="14" height="14" viewBox="0 0 16 16" fill="none">
                    <path d="M2 2h8l4 4v8H2V2z" fill="#ffd54f" stroke="#b8860b" strokeWidth="1"/>
                    <path d="M10 2v4h4" fill="none" stroke="#b8860b" strokeWidth="1"/>
                    <rect x="4" y="8" width="8" height="1" fill="#b8860b"/>
                    <rect x="4" y="10" width="8" height="1" fill="#b8860b"/>
                    <rect x="4" y="12" width="5" height="1" fill="#b8860b"/>
                  </svg>
                  Browse...
                </span>
              </label>
              
              {file && (
                <div className="flex-1 xp-input flex items-center px-2">
                  <svg width="12" height="12" viewBox="0 0 16 16" className="mr-2" fill="none">
                    <rect x="2" y="3" width="12" height="10" rx="1" fill="#217346"/>
                    <text x="8" y="11" fontSize="8" fill="white" textAnchor="middle" fontWeight="bold">X</text>
                  </svg>
                  <span className="text-xs truncate">{file.name}</span>
                </div>
              )}
            </div>
          </div>

          {file && (
            <div className="flex gap-2 mt-3">
              <button
                onClick={convertToICS}
                disabled={status.type === 'loading'}
                className="xp-button"
              >
                {status.type === 'loading' ? 'Converting...' : 'Convert to Calendar'}
              </button>
              <button
                onClick={reset}
                disabled={status.type === 'loading'}
                className="xp-button"
              >
                Clear
              </button>
            </div>
          )}
        </div>

        {/* Progress indicator */}
        {status.type === 'loading' && (
          <div className="xp-messagebox mb-4">
            <div className="flex items-center gap-3 mb-3">
              <div className="animate-spin" style={{ fontSize: '24px' }}>⏳</div>
              <div>
                <p className="text-xs font-bold mb-1">Please wait...</p>
                <p className="text-xs">Converting your spreadsheet to calendar format.</p>
              </div>
            </div>
            <div className="xp-progress">
              <div className="xp-progress-bar" style={{ width: '100%' }} />
            </div>
          </div>
        )}

        {/* Success message */}
        {status.type === 'success' && (
          <div className="xp-messagebox mb-4">
            <div className="flex items-start gap-3">
              <div style={{ fontSize: '32px', lineHeight: '1' }}>
                <svg width="32" height="32" viewBox="0 0 32 32" fill="none">
                  <circle cx="16" cy="16" r="15" fill="#4caf50" stroke="#2e7d32" strokeWidth="2"/>
                  <path d="M9 16l5 5l9-9" stroke="white" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round"/>
                </svg>
              </div>
              <div className="flex-1">
                <p className="text-sm font-bold mb-1">Operation Completed Successfully</p>
                <p className="text-xs" style={{ lineHeight: '1.5' }}>
                  {status.message}
                </p>
                <button onClick={() => setStatus({ type: 'idle' })} className="xp-button mt-3">
                  OK
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Error message */}
        {status.type === 'error' && (
          <div className="xp-messagebox mb-4">
            <div className="flex items-start gap-3">
              <div style={{ fontSize: '32px', lineHeight: '1' }}>
                <svg width="32" height="32" viewBox="0 0 32 32" fill="none">
                  <circle cx="16" cy="16" r="15" fill="#ff5252" stroke="#c62828" strokeWidth="2"/>
                  <path d="M12 12l8 8M20 12l-8 8" stroke="white" strokeWidth="3" strokeLinecap="round"/>
                </svg>
              </div>
              <div className="flex-1">
                <p className="text-sm font-bold mb-1">Error</p>
                <p className="text-xs" style={{ lineHeight: '1.5' }}>
                  {status.message}
                </p>
                <button onClick={() => setStatus({ type: 'idle' })} className="xp-button mt-3">
                  OK
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Instructions groupbox */}
        <div className="xp-groupbox">
          <div className="xp-groupbox-title">Instructions</div>
          
          <div className="text-xs" style={{ lineHeight: '1.6', color: '#000' }}>
            <p className="mb-2 font-bold">How to use this converter:</p>
            <ol className="list-decimal ml-5 space-y-1 mb-3">
              <li>Click <strong>Browse...</strong> to select your Excel file</li>
              <li>Click <strong>Convert to Calendar</strong> to process the file</li>
              <li>The .ics file will download automatically</li>
              <li>Import the .ics file into your calendar application</li>
            </ol>
            
            <p className="mb-2 font-bold">Expected Excel format:</p>
            <div className="bg-white border border-gray-400 p-2 text-xs">
              <p className="mb-1">Your spreadsheet should contain these columns:</p>
              <ul className="list-disc ml-5 space-y-1">
                <li><strong>title</strong> or <strong>event</strong> - Event name</li>
                <li><strong>start</strong> or <strong>date</strong> - Start date/time</li>
                <li><strong>end</strong> - End date/time (optional)</li>
                <li><strong>description</strong> - Event details (optional)</li>
                <li><strong>location</strong> - Event location (optional)</li>
              </ul>
            </div>
          </div>
        </div>

        {/* Status bar */}
        <div style={{
          marginTop: '16px',
          padding: '3px 6px',
          background: '#ece9d8',
          border: '1px solid',
          borderColor: '#808080 #ffffff #ffffff #808080',
          fontSize: '10px',
          color: '#000',
          display: 'flex',
          alignItems: 'center',
          gap: '8px',
          boxShadow: 'inset 1px 1px 0 rgba(0,0,0,0.1)'
        }}>
          <svg width="12" height="12" viewBox="0 0 16 16" fill="none">
            <circle cx="8" cy="8" r="7" fill="#4caf50"/>
            <path d="M8 4v5l3 2" stroke="white" strokeWidth="1.5" strokeLinecap="round"/>
          </svg>
          <span>Ready</span>
        </div>
      </div>
    </div>
  );
}