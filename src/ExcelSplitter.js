import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';
import { Container, TextField, Button, Typography, Grid, CircularProgress, Backdrop, Card, CardContent } from '@mui/material';

const ExcelSplitter = () => {
    const [file, setFile] = useState(null);
    const [chunkSize, setChunkSize] = useState(100); // Default: 100 rows per file
    const [isLoading, setIsLoading] = useState(false); // Loading state
    const [fileName, setFileName] = useState(''); // Storing file name

    const handleFileChange = (e) => {
        const selectedFile = e.target.files[0];
        if (selectedFile) {
            setFile(selectedFile);
            setFileName(selectedFile.name); // Set file name when a file is selected
        }
    };

    const handleChunkSizeChange = (e) => {
        setChunkSize(e.target.value);
    };

    const splitExcelFileAndZip = async () => {
        if (!file) {
            alert("Please select a file first!");
            return;
        }

        if (chunkSize <= 0) {
            alert("Chunk size must be greater than 0!");
            return;
        }

        // Set loading to true when process starts
        setIsLoading(true);

        const reader = new FileReader();

        reader.onload = async (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            if (!workbook || !workbook.SheetNames.length) {
                alert("Invalid or empty file.");
                setIsLoading(false); // Set loading to false in case of error
                return;
            }

            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }); // Get data with header

            const totalRows = jsonData.length;
            const header = jsonData[0]; // Get header row (first row)
            const bodyData = jsonData.slice(1); // Get data without header

            const numChunks = Math.ceil(bodyData.length / chunkSize); // Calculate number of files to create

            const zip = new JSZip();
            const originalFileName = file.name.replace(/\.[^/.]+$/, ""); // Remove extension from file name

            for (let i = 0; i < numChunks; i++) {
                const startRow = i * chunkSize;
                const endRow = startRow + chunkSize;
                const chunkData = bodyData.slice(startRow, endRow);

                // Add header back to each chunk
                const chunkWithHeader = [header, ...chunkData];

                const newWorkbook = XLSX.utils.book_new();
                const newWorksheet = XLSX.utils.aoa_to_sheet(chunkWithHeader);
                XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, `Sheet1`);

                // Convert each chunk into binary and add it to the zip file
                const wbout = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' });
                zip.file(`${originalFileName}_part_${i + 1}.xlsx`, wbout);
            }

            // Generate the zip and save it
            zip.generateAsync({ type: 'blob' }).then((content) => {
                saveAs(content, `${originalFileName}_chunks.zip`);
                setIsLoading(false); // Set loading to false after process finishes
            });
        };

        reader.readAsArrayBuffer(file);
    };

    return (
        <Container maxWidth="sm" style={{ marginTop: '50px', backgroundColor: '#F0ECE5',boxShadow: '0 8px 20px rgba(0, 0, 0, 0.2)', padding: '20px', borderRadius: '8px' }}>
            <Typography
                variant="h4"
                gutterBottom
                align="center"
                style={{
                    fontWeight: 'bold',       // Membuat teks menjadi bold
                    textShadow: '2px 2px 5px rgba(0, 0, 0, 0.3)',  // Menambahkan efek shadow
                    color: '#31304D',         // Warna teks
                }}
            >
                Excel Splitter
            </Typography>


            <Card style={{backgroundColor: '#FFFFFF' }}>
                <CardContent>
                    <Grid container spacing={3}>
                        <Grid item xs={12}>
                            <input
                                type="file"
                                onChange={handleFileChange}
                                accept=".xlsx,.xls,.csv"
                                style={{ display: 'none' }}
                                id="upload-file"
                            />
                            <label htmlFor="upload-file">
                                <Button variant="contained" component="span" fullWidth disabled={isLoading} style={{ backgroundColor: '#31304D', color: '#FFFFFF' }}>
                                    Choose Excel/CSV File
                                </Button>
                            </label>
                        </Grid>
                        {fileName && (
                            <Grid item xs={12}>
                                <Typography variant="body1" color="textSecondary">
                                    Selected File: <strong>{fileName}</strong>
                                </Typography>
                            </Grid>
                        )}
                        <Grid item xs={12}>
                            <TextField
                                label="Rows per File"
                                type="number"
                                value={chunkSize}
                                onChange={handleChunkSizeChange}
                                fullWidth
                                disabled={isLoading} // Disable input during loading
                            />
                        </Grid>
                        <Grid item xs={12}>
                            <Button
                                variant="contained"
                                color="primary"
                                onClick={splitExcelFileAndZip}
                                fullWidth
                                disabled={isLoading} // Disable button during loading
                                style={{ backgroundColor: '#31304D', color: '#FFFFFF' }}
                            >
                                {isLoading ? "Processing..." : "Split and Download as ZIP"}
                            </Button>
                        </Grid>
                    </Grid>
                </CardContent>
            </Card>

            {/* Loading Animation */}
            <Backdrop style={{ zIndex: 1200 }} open={isLoading}>
                <CircularProgress size={60} thickness={5} />
            </Backdrop>

            {/* Developer Information */}
            <Typography variant="body2" color="textSecondary" align="center" style={{ marginTop: '30px' }}>
                Created by <strong>VICI</strong>. Current version: <strong>0.1.2</strong>.
                Latest update: <em>Supports splitting data files in <strong>XLSX, XLS, and CSV</strong> formats</em>.
            </Typography>
        </Container>
    );
};

export default ExcelSplitter;
