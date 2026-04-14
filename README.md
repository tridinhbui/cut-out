# Smithfield FPork Total Cut-Out Extractor

A Next.js web application for intelligently extracting Total Cut-Out Margin (Total SFI) data from Smithfield FPork financial Excel reports.

## Features

- **Drag & Drop Upload**: Support for multiple Excel (.xlsx) files
- **Intelligent Extraction**: Auto-detects data columns and extracts values with timestamps
- **Time Period Detection**: Identifies monthly and weekly periods with dates and week codes
- **History Tracking**: Saves all extraction results locally for review
- **Beautiful UI**: Orange and blue themed interface
- **Fact Checking**: Clear display of extracted data for verification

## Extraction Logic

1. **Row Detection**: Finds "Total Cut-Out Margin" row in column A (around row 131)
2. **Column Auto-Detection**: Scans headers for "Total SFI" (row 2-3) and "Overall" (row 6)
3. **Date Extraction**:
   - Monthly: Reads date from B4 (serial) and month name from B5
   - Weekly: Scans left for "wk ending" label, extracts serial date and week code
4. **Value Extraction**: Pulls numeric values from detected columns at target row

## Setup

1. Install dependencies: `npm install`
2. Run the development server: `npm run dev`
3. Open [http://localhost:3000](http://localhost:3000)
4. Drag and drop Excel files to extract data

## Usage

- Upload multiple files at once
- View latest results and historical extractions
- Data is stored locally in browser storage
- Clear history when needed

## Technology Stack

- Next.js 16
- TypeScript
- Tailwind CSS
- XLSX library for Excel processing

## Copyright

© 2026 Tri Bui Team - Corporate Finance FP&A. All rights reserved.
