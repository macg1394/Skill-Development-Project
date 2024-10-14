document.addEventListener("DOMContentLoaded", () => {
    const table = document.querySelector(".table");
    const isUploadPage = document.getElementById('uploadPdf') !== null;

    if (!isUploadPage) {
        // Load initial data from Sheet1.xlsx on the homepage
        (async () => {
            try {
                const response = await fetch("./Sheet1.xlsx");
                console.log("Fetch response:", response);

                if (!response.ok) {
                    throw new Error(`Failed to fetch: ${response.statusText}`);
                }
                const workbook = XLSX.read(await response.arrayBuffer());
                const worksheetNames = workbook.SheetNames;

                if (table) {
                    worksheetNames.forEach(name => {
                        const html = XLSX.utils.sheet_to_html(workbook.Sheets[name]);
                        table.innerHTML += `<h1>${name}</h1>${html}`;
                    });

                    filterRows(); // Apply filters initially
                } else {
                    console.error("Table element not found in the DOM");
                    alert("Error: Table element not found. Please check your HTML structure.");
                }
            } catch (error) {
                console.error("Error loading Sheet1.xlsx:", error);
                alert("Error loading initial data. Check the console for details.");
            }
        })();

        // Add event listeners to custom checkboxes
        document.querySelectorAll('.dropdown-content input').forEach(checkbox => {
            checkbox.addEventListener('change', filterRows);
        });
    } else {
        // Event listener for PDF upload
        const processPdfButton = document.getElementById('processPdf');
        if (processPdfButton) {
            processPdfButton.addEventListener('click', () => {
                const pdfFileInput = document.getElementById('uploadPdf');
                if (pdfFileInput) {
                    const pdfFile = pdfFileInput.files[0];
                    if (pdfFile) {
                        console.log('PDF file selected:', pdfFile.name);
                        processPdf(pdfFile);
                    } else {
                        alert('Please upload a PDF file before clicking Extract Data.');
                    }
                } else {
                    console.error("Upload PDF input not found");
                    alert("Error: PDF upload input not found. Please check your HTML structure.");
                }
            });
        } else {
            console.error("Process PDF button not found");
        }
    }

    // Function to extract data from uploaded PDF
    async function processPdf(pdfFile) {
        try {
            const pdfjsLib = window['pdfjs-dist/build/pdf'];
            if (!pdfjsLib) {
                throw new Error("PDF.js library not found. Make sure it's properly included.");
            }
            pdfjsLib.GlobalWorkerOptions.workerSrc = '//mozilla.github.io/pdf.js/build/pdf.worker.js';

            const reader = new FileReader();

            reader.onload = async function () {
                const typedArray = new Uint8Array(reader.result);
                const pdf = await pdfjsLib.getDocument(typedArray).promise;
                let pdfExtractedText = '';

                for (let i = 1; i <= pdf.numPages; i++) {
                    const page = await pdf.getPage(i);
                    const textContent = await page.getTextContent();
                    textContent.items.forEach(item => {
                        pdfExtractedText += item.str + ' ';
                    });
                }

                console.log('Extracted PDF Text:', pdfExtractedText);
                alert("PDF data extraction complete! Check the console for extracted text.");
                await appendDataToExcel(pdfExtractedText);
            };

            reader.readAsArrayBuffer(pdfFile);
        } catch (error) {
            console.error("Error processing the PDF:", error);
            alert("Error processing the PDF. Check the console for details.");
        }
    }

    // Function to append extracted data to Competitive Exam Data.xlsx
    async function appendDataToExcel(extractedText) {
        try {
            const response = await fetch("./Competitive Exam Data.xlsx");
            if (!response.ok) {
                throw new Error(`Failed to fetch: ${response.statusText}`);
            }
    
            const existingWorkbook = XLSX.read(await response.arrayBuffer());
            const sheetName = existingWorkbook.SheetNames[0];
            const worksheet = existingWorkbook.Sheets[sheetName];
    
            // Split the extracted text into rows and cells properly
            const newData = extractedText.trim().split(/\n+/).map(line => line.trim().split(/\s+/));
    
            // Ensure the worksheet has a valid range, or create a new one
            let range = worksheet['!ref'] ? XLSX.utils.decode_range(worksheet['!ref']) : { s: { r: 0, c: 0 }, e: { r: 0, c: newData[0].length - 1 } };
    
            const startRow = range.e.r + 1;
    
            newData.forEach((row, index) => {
                row.forEach((cell, cellIndex) => {
                    const cellRef = XLSX.utils.encode_cell({ r: startRow + index, c: cellIndex });
                    worksheet[cellRef] = { v: cell };
                });
            });
    
            // Update the worksheet range
            worksheet['!ref'] = XLSX.utils.encode_range({
                s: { r: 0, c: 0 },
                e: { r: startRow + newData.length - 1, c: newData[0].length - 1 }
            });
    
            // Generate the updated Excel file and prompt download
            const updatedWorkbookBlob = new Blob([XLSX.write(existingWorkbook, { bookType: 'xlsx', type: 'array' })], { type: 'application/octet-stream' });
            const url = URL.createObjectURL(updatedWorkbookBlob);
    
            const a = document.createElement('a');
            a.href = url;
            a.download = 'Competitive Exam Data.xlsx';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
    
            alert("Data has been appended to Competitive Exam Data.xlsx and the file has been downloaded.");
        } catch (error) {
            console.error("Error appending data to Excel:", error);
            alert("Error appending data to Excel. Check the console for details.");
        }
    }
    
    
    // Function to get selected (checked) values from a dropdown
    function getCheckedValues(containerId) {
        const container = document.getElementById(containerId);
        if (!container) {
            console.error(`Container with id '${containerId}' not found`);
            return [];
        }
        return [...container.querySelectorAll('input:checked')]
            .map(cb => cb.value.trim().toLowerCase());
    }

    // Function to filter rows based on selected filters
    function filterRows() {
        const examFilterValues = getCheckedValues("examDropdown");
        const branchFilterValues = getCheckedValues("branchDropdown");
        const batchFilterValues = getCheckedValues("batchDropdown");
        const rows = document.querySelectorAll(".table table tbody tr");

        rows.forEach(row => {
            const examCell = row.querySelector("td:nth-child(3)");
            const branchCell = row.querySelector("td:nth-child(4)");
            const batchCell = row.querySelector("td:nth-child(6)");

            const examCellText = examCell ? examCell.textContent.trim().toLowerCase() : '';
            const branchCellText = branchCell ? branchCell.textContent.trim().toLowerCase() : '';
            const batchCellText = batchCell ? batchCell.textContent.trim().toLowerCase() : '';

            const showRow = 
                (examFilterValues.length === 0 || examFilterValues.includes(examCellText)) &&
                (branchFilterValues.length === 0 || branchFilterValues.includes(branchCellText)) &&
                (batchFilterValues.length === 0 || batchFilterValues.includes(batchCellText));

            row.style.display = showRow ? "" : "none";
        });
    }
});