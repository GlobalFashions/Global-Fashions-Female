document.getElementById('processButton').addEventListener('click', () => {
    const fileInput = document.getElementById('fileInput');
    
    // Ensure a file is selected
    if (!fileInput.files.length) {
        alert('Please select a file first.');
        return;
    }

    const file = fileInput.files[0];
    const reader = new FileReader();

    reader.onload = (event) => {
        try {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            // Assuming the source data is in the first sheet
            const sourceSheet = workbook.Sheets[workbook.SheetNames[0]];
            const sourceData = XLSX.utils.sheet_to_json(sourceSheet);

            // Process data to match the Excel sheet's column order
            const formattedData = sourceData.map(item => {
                const bodyFabric = item['Body/Fabric'] || '';
                const handleTitle = bodyFabric.replace(/\s+/g, '').toLowerCase();
                const tags = bodyFabric.replace(/\s+/g, '');

                return {
                    Handle: handleTitle,
                    Command: 'MERGE', // Adding static command for column 'Command'
                    Title: handleTitle,
                    'Body HTML': item['Material Composition by %'] || '',
                    Vendor: item['Division Name'] || '',
                    Type: item['Category'] || '',
                    Tags: handleTitle,
                    'Tags Command': 'REPLACE', // Adding static command for 'Tags Command'
                    Status: 'active', // Static value for 'Status'
                    'Total Inventory Qty': item['Total Inventory Qty'] || 1,
                    'Image Src': '', // Placeholder as image is not part of your original data
                    'Image Command': '', // Placeholder for 'Image Command'
                    'Option1 Name': 'Color',
                    'Option1 Value': item['Color Description'] || '',
                    'Option2 Name': 'Size',
                    'Option2 Value': item['Size'] || '',
                    'Option3 Name': '', // Assuming no Option3
                    'Option3 Value': '', // Assuming no Option3
                    'Variant SKU': item['EAN'] || '',
                    'Variant Barcode': item['EAN'] || '',
                    'Variant Weight': item['Shipped Gross'] || '',
                    'Variant Weight Unit': 'g',
                    'Variant Price': item['Cost'] || '',
                    'Variant Compare At Price': item['Cost'] || '',
                    'Variant Taxable': 'TRUE',
                    'Variant Inventory Tracker': 'shopify',
                    'Variant Inventory Policy': 'deny',
                    'Variant Fulfillment Service': 'manual',
                    'Variant Inventory Qty': item['Total Inventory Qty'] || 1,
                    'Variant Cost': item['Cost'] || '',
                    'Variant Country of Origin': 'Italy', // Placeholder if data is not provided
                    'Variant Metafield: mm-google-shopping.age_group [single_line_text_field]': 'Adult',
                    'Variant Metafield: mm-google-shopping.condition [single_line_text_field]': 'New',
                    'Variant Metafield: mm-google-shopping.gender [single_line_text_field]': 'Female',
                    'Variant Metafield: mm-google-shopping.custom_label_1 [single_line_text_field]': item['Size'] || '',
                    'Variant Metafield: mm-google-shopping.custom_label_2 [single_line_text_field]': item['Color Description'] || ''
                };
            });

            // Convert formatted data to a CSV string
            const csv = XLSX.utils.sheet_to_csv(XLSX.utils.json_to_sheet(formattedData));

            // Create a downloadable file
            const blob = new Blob([csv], { type: 'text/csv' });

            // Create an object URL and set it to the download link
            const downloadLink = document.getElementById('downloadLink');
            downloadLink.href = URL.createObjectURL(blob);
            downloadLink.download = 'formatted_inventory.csv';
            downloadLink.style.display = 'block'; // Show link after file is processed
            downloadLink.innerHTML = 'Download Formatted CSV File';
        } catch (error) {
            alert('An error occurred while processing the file.');
            console.error(error);
        }
    };

    reader.readAsArrayBuffer(file);
});

// Hide the download link by default
document.getElementById('downloadLink').style.display = 'none';
