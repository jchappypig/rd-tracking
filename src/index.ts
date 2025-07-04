import ExcelJS from 'exceljs';
import path from 'path';

type Ticket = {
    [key: string]: any;
    'Work item key': string;
    'Parent'?: string;
    'Linked work items'?: string;
    'R&DTI Activity'?: string;
    'Work type'?: string;
    'Work Hours in Progress'?: string;
    'Sum of WIP hours'?: string;
};

const INPUT_FILE = 'input.xlsx';
const OUTPUT_FILE = 'output_with_rdti.xlsx';

async function main() {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(path.resolve(INPUT_FILE));
        const sheet = workbook.getWorksheet(1); // Use first sheet

        if (!sheet) {
            console.error('❌ Sheet not found.');
            return;
        }

        const headers: { [key: string]: number } = {};
        sheet.getRow(1).eachCell((cell, colNumber) => {
            if (cell.value) headers[cell.value.toString()] = colNumber;
        });

        const rows: Ticket[] = [];
        const ticketMap: Record<string, Ticket> = {};

        sheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return; // skip header

            const rowData: Ticket = {} as Ticket;
            for (const [header, col] of Object.entries(headers)) {
                const cellValue = row.getCell(col).value;
                // Handle different cell value types properly
                if (cellValue === null || cellValue === undefined) {
                    rowData[header] = '';
                } else if (typeof cellValue === 'object' && 'text' in cellValue) {
                    // Handle rich text objects
                    rowData[header] = cellValue.text;
                } else if (typeof cellValue === 'object' && 'hyperlink' in cellValue) {
                    // Handle hyperlink objects
                    rowData[header] = cellValue.hyperlink;
                } else {
                    rowData[header] = cellValue.toString();
                }
            }

            rows.push(rowData);
            ticketMap[rowData['Work item key']] = rowData;
        });

        // Helper function to check if a ticket is a MAP ticket
        function isMapTicket(ticket: Ticket): boolean {
            // MAP tickets are "Idea" type tickets with work item keys starting with "MAP-"
            return ticket['Work type'] === 'Idea' && ticket['Work item key'].startsWith('MAP-');
        }

        // Helper function to parse linked work items
        function parseLinkedWorkItems(linkedWorkItems: string): string[] {
            if (!linkedWorkItems) return [];
            return linkedWorkItems.split(',').map(item => item.trim()).filter(Boolean);
        }

        // Helper function to find MAP tickets in linked work items
        function findMapTicketsInLinked(ticket: Ticket): Ticket[] {
            const linkedItems = parseLinkedWorkItems(ticket['Linked work items'] || '');
            return linkedItems
                .map(itemKey => ticketMap[itemKey])
                .filter(linkedTicket => linkedTicket && isMapTicket(linkedTicket));
        }

        // Helper function to get R&DTI Activity from linked MAP tickets
        function getRDTIFromLinkedMap(ticket: Ticket): string | null {
            const mapTickets = findMapTicketsInLinked(ticket);
            for (const mapTicket of mapTickets) {
                if (mapTicket['R&DTI Activity']) {
                    return mapTicket['R&DTI Activity'];
                }
            }
            return null;
        }

        // Helper function to get all children of a ticket
        function getChildren(parentTicket: Ticket): Ticket[] {
            return rows.filter(row => 
                row['Parent'] && 
                row['Parent'].startsWith(parentTicket['Work item key'] + ' ')
            );
        }

        // Helper function to recursively get all descendants
        function getAllDescendants(ticket: Ticket): Ticket[] {
            const children = getChildren(ticket);
            let descendants = [...children];
            
            for (const child of children) {
                descendants = descendants.concat(getAllDescendants(child));
            }
            
            return descendants;
        }

        // Helper function to parse time duration to hours
        function parseTimeToHours(timeStr: string): number {
            if (!timeStr || timeStr === '') return 0;
            
            // Handle different time formats like "1d 2h 30m", "5h 15m", "30m", etc.
            let totalHours = 0;
            const timeRegex = /(\d+(?:\.\d+)?)\s*([dhm])/g;
            let match;
            
            while ((match = timeRegex.exec(timeStr)) !== null) {
                const value = parseFloat(match[1]);
                const unit = match[2];
                
                switch (unit) {
                    case 'd': totalHours += value * 8; break; // Assuming 8 hours per day
                    case 'h': totalHours += value; break;
                    case 'm': totalHours += value / 60; break;
                }
            }
            
            return totalHours;
        }

        // Helper function to get WIP hours for a ticket and its children
        function getWIPHours(ticket: Ticket): number {
            const ticketHours = parseTimeToHours(ticket['Work Hours in Progress'] || '');
            const children = getChildren(ticket);
            
            if (children.length === 0) {
                return ticketHours;
            }
            
            // Calculate sum of children's WIP hours
            const childrenSum = children.reduce((sum, child) => sum + getWIPHours(child), 0);
            
            // Return the larger of ticket's own hours or sum of children's hours
            return Math.max(ticketHours, childrenSum);
        }

        // Process tickets with new logic
        let processedCount = 0;
        const processedTickets = new Set<string>();

        for (const ticket of rows) {
            // Skip if already processed
            if (processedTickets.has(ticket['Work item key'])) continue;
            
            // Skip if ticket already has R&DTI Activity
            if (ticket['R&DTI Activity']) continue;
            
            // If ticket is not a MAP ticket, look for R&DTI Activity in linked MAP tickets
            if (!isMapTicket(ticket)) {
                const rdtiFromLinkedMap = getRDTIFromLinkedMap(ticket);
                
                if (rdtiFromLinkedMap) {
                    // Set R&DTI Activity for this ticket
                    ticket['R&DTI Activity'] = rdtiFromLinkedMap;
                    processedTickets.add(ticket['Work item key']);
                    processedCount++;
                    
                    // Propagate to all descendants
                    const descendants = getAllDescendants(ticket);
                    for (const descendant of descendants) {
                        if (!descendant['R&DTI Activity']) {
                            descendant['R&DTI Activity'] = rdtiFromLinkedMap;
                            processedTickets.add(descendant['Work item key']);
                            processedCount++;
                        }
                    }
                }
            }
        }

        // Calculate Sum of WIP hours for MAP tickets with R&DTI Activities
        let mapTicketsProcessed = 0;
        
        for (const ticket of rows) {
            // Initialize Sum of WIP hours column
            ticket['Sum of WIP hours'] = '';
            
            // Only process MAP tickets with R&DTI Activities
            if (isMapTicket(ticket) && ticket['R&DTI Activity']) {
                const linkedItems = parseLinkedWorkItems(ticket['Linked work items'] || '');
                let totalWIPHours = 0;
                
                if (linkedItems.length > 0) {
                    // Track which tickets we've already counted to avoid double counting
                    const countedTickets = new Set<string>();
                    
                    // Helper function to collect all descendants of a ticket
                    function collectAllDescendants(parentTicket: Ticket): Set<string> {
                        const descendants = new Set<string>();
                        const children = getChildren(parentTicket);
                        
                        for (const child of children) {
                            descendants.add(child['Work item key']);
                            // Recursively add grandchildren, etc.
                            const grandchildren = collectAllDescendants(child);
                            grandchildren.forEach(gchild => descendants.add(gchild));
                        }
                        
                        return descendants;
                    }
                    
                    // First pass: collect all descendants of all linked items to identify potential double counts
                    const allDescendants = new Set<string>();
                    for (const linkedId of linkedItems) {
                        const linkedTicket = ticketMap[linkedId];
                        if (linkedTicket && !isMapTicket(linkedTicket)) {
                            const descendants = collectAllDescendants(linkedTicket);
                            descendants.forEach(desc => allDescendants.add(desc));
                        }
                    }
                    
                    // Second pass: calculate WIP hours, avoiding double counting
                    for (const linkedId of linkedItems) {
                        const linkedTicket = ticketMap[linkedId];
                        if (!linkedTicket) continue;
                        
                        // Skip if this ticket is a descendant of another linked item (to avoid double counting)
                        if (allDescendants.has(linkedId)) {
                            continue;
                        }
                        
                        if (!isMapTicket(linkedTicket)) {
                            // Non-MAP ticket: get WIP hours (considering children)
                            const wipHours = getWIPHours(linkedTicket);
                            totalWIPHours += wipHours;
                            countedTickets.add(linkedId);
                        } else {
                            // MAP ticket: only include if it doesn't have R&DTI Activity
                            if (!linkedTicket['R&DTI Activity']) {
                                const wipHours = getWIPHours(linkedTicket);
                                totalWIPHours += wipHours;
                                countedTickets.add(linkedId);
                            }
                        }
                    }
                }
                
                // Set the sum of WIP hours (rounded to 2 decimal places)
                ticket['Sum of WIP hours'] = totalWIPHours.toFixed(2);
                mapTicketsProcessed++;
            }
        }
        
        console.log(`✅ Processed ${mapTicketsProcessed} MAP tickets with R&DTI Activities for WIP hours calculation`);
           
           // Write output
           const outWorkbook = new ExcelJS.Workbook();
           const outSheet = outWorkbook.addWorksheet('Results');
  
           // Make sure to include the new "Sum of WIP hours" column
           const columnOrder = [...Object.keys(headers), 'Sum of WIP hours'];
           outSheet.addRow(columnOrder); // header
  
           rows.forEach(row => {
               const values = columnOrder.map(key => row[key] || '');
               outSheet.addRow(values);
           });

           // Create new sheet with transformed data
           const transformedSheet = outWorkbook.addWorksheet('Transformed Data');
           
           // Add headers for the new sheet
           const transformedHeaders = ['Project', 'Who', 'Role', 'Activity Type', 'Hours/Cost', 'Phase', 'Work Item'];
           transformedSheet.addRow(transformedHeaders);
           
           // Filter rows that have R&DTI Activity and transform them
           const rowsWithRDTI = rows.filter(row => row['R&DTI Activity']);
           
           rowsWithRDTI.forEach(row => {
               // Column 1: R&DTI Activity → Project
               const project = row['R&DTI Activity'] || '';
               
               // Column 2: Assignee or Mob (whichever has values) → Who
               const who = row['Assignee'] || row['Mob'] || '';
               
               // Column 3: Role → "Employee" (static value)
               const role = 'Employee';
               
               // Column 4: Activity Type → "Core" (except if R&DTI Activity is "Platform", then "Support")
               const activityType = row['R&DTI Activity'] === 'Platform' ? 'Support' : 'Core';
               
               // Column 5: Sum of WIP hours → Hours/Cost
               const hoursCost = row['Sum of WIP hours'] || '';
               
               // Column 6: Phase → "Development" (static value)
               const phase = 'Development';
               
               // Column 7: Work Item
               const workItem = row['Work item key'] || '';
               
               transformedSheet.addRow([project, who, role, activityType, hoursCost, phase, workItem]);
           });
  
           await outWorkbook.xlsx.writeFile(path.resolve(OUTPUT_FILE));
           console.log(`✅ File saved to ${OUTPUT_FILE}`);
           console.log(`✅ Updated ${processedCount} rows with R&DTI Activity from linked MAP tickets`);
           console.log(`✅ Created transformed data sheet with ${rowsWithRDTI.length} rows`);
           
       } catch (error) {
           console.error('❌ Error occurred:', error);
           throw error;
       }
   }
   
   main().catch(err => console.error('❌ Error:', err));