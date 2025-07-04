import ExcelJS from 'exceljs';
import path from 'path';

type Ticket = {
    [key: string]: any;
    'Work item key': string;
    'Parent'?: string;
    'Linked work items'?: string;
    'R&DTI Activity'?: string;
    'Work type'?: string;
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
            return rows.filter(row => row['Parent'] === parentTicket['Work item key']);
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

        // Write output
        const outWorkbook = new ExcelJS.Workbook();
        const outSheet = outWorkbook.addWorksheet('Results');

        const columnOrder = Object.keys(headers);
        outSheet.addRow(columnOrder); // header

        rows.forEach(row => {
            const values = columnOrder.map(key => row[key] || '');
            outSheet.addRow(values);
        });

        await outWorkbook.xlsx.writeFile(path.resolve(OUTPUT_FILE));
        console.log(`✅ File saved to ${OUTPUT_FILE}`);
        console.log(`✅ Updated ${processedCount} rows with R&DTI Activity from linked MAP tickets`);
        
    } catch (error) {
        console.error('❌ Error occurred:', error);
        throw error;
    }
}

main().catch(err => console.error('❌ Error:', err));
