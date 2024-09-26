import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';
import { Container, TextField, Button, Typography, Grid, CircularProgress, Backdrop } from '@mui/material';

const ExcelSplitter = () => {
  const [file, setFile] = useState(null);
  const [chunkSize, setChunkSize] = useState(100); // Default: 100 baris per file
  const [isLoading, setIsLoading] = useState(false); // State untuk loading
  const [fileName, setFileName] = useState(''); // Menyimpan nama file

  const handleFileChange = (e) => {
    const selectedFile = e.target.files[0];
    if (selectedFile) {
      setFile(selectedFile);
      setFileName(selectedFile.name); // Set nama file yang dipilih
    }
  };

  const handleChunkSizeChange = (e) => {
    setChunkSize(e.target.value);
  };

  const splitExcelFileAndZip = async () => {
    if (!file) {
      alert("Pilih file terlebih dahulu!");
      return;
    }

    if (chunkSize <= 0) {
      alert("Ukuran chunk harus lebih dari 0!");
      return;
    }

    // Set loading menjadi true saat mulai proses
    setIsLoading(true);

    const reader = new FileReader();

    reader.onload = async (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      if (!workbook || !workbook.SheetNames.length) {
        alert("File tidak valid atau kosong.");
        setIsLoading(false); // Set loading menjadi false jika ada error
        return;
      }

      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }); // Ambil data dengan header

      const totalRows = jsonData.length;
      const header = jsonData[0]; // Ambil baris header (baris pertama)
      const bodyData = jsonData.slice(1); // Ambil data tanpa header

      const numChunks = Math.ceil(bodyData.length / chunkSize); // Hitung jumlah file yang akan dibuat

      const zip = new JSZip();
      const originalFileName = file.name.replace(/\.[^/.]+$/, ""); // Hapus ekstensi dari nama file

      for (let i = 0; i < numChunks; i++) {
        const startRow = i * chunkSize;
        const endRow = startRow + chunkSize;
        const chunkData = bodyData.slice(startRow, endRow);

        // Tambahkan header kembali ke setiap chunk
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
        setIsLoading(false); // Set loading menjadi false setelah proses selesai
      });
    };

    reader.readAsArrayBuffer(file);
  };

  return (
    <Container maxWidth="sm" style={{ marginTop: '50px' }}>
      <Typography variant="h4" gutterBottom align="center">
        Excel Splitter
      </Typography>
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
            <Button variant="contained" component="span" fullWidth disabled={isLoading}>
              Pilih File Excel/CSV
            </Button>
          </label>
        </Grid>
        {fileName && (
          <Grid item xs={12}>
            <Typography variant="body1" color="textSecondary">
              File yang dipilih: <strong>{fileName}</strong>
            </Typography>
          </Grid>
        )}
        <Grid item xs={12}>
          <TextField
            label="Jumlah Baris per File"
            type="number"
            value={chunkSize}
            onChange={handleChunkSizeChange}
            fullWidth
            disabled={isLoading} // Disable input saat loading
          />
        </Grid>
        <Grid item xs={12}>
          <Button
            variant="contained"
            color="primary"
            onClick={splitExcelFileAndZip}
            fullWidth
            disabled={isLoading} // Disable tombol saat loading
          >
            {isLoading ? "Memproses..." : "Split dan Unduh sebagai ZIP"}
          </Button>
        </Grid>
      </Grid>

      {/* Animasi Loading */}
      <Backdrop style={{ zIndex: 1200 }} open={isLoading}>
        <CircularProgress size={60} thickness={5} />
      </Backdrop>

      {/* Keterangan Developer */}
      <Typography variant="body2" color="textSecondary" align="center" style={{ marginTop: '30px' }}>
        Dibuat oleh <strong>VICi</strong> - Software Development House. Versi saat ini: <strong>0.1.2</strong>. 
        Update terbaru: <em>Support untuk memecah file data dengan format <strong>XLSX, XLS, dan CSV</strong></em>.
      </Typography>
    </Container>
  );
};

export default ExcelSplitter;
