import { NextResponse } from 'next/server';
import * as XLSX from 'xlsx';
import path from 'path';
import fs from 'fs';

// Mengarah ke file Data.xlsx di folder utama project
const filePath = path.join(process.cwd(), 'Data.xlsx'); 

export async function GET() {
  try {
    if (!fs.existsSync(filePath)) {
      return NextResponse.json([]);
    }
    const fileBuffer = fs.readFileSync(filePath);
    const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
    
    // Konversi format Excel ke format Dashboard
    const formattedData = data.map((row, index) => ({
      id: index + 1,
      alertType: row['Alert Type'] || '-',
      shift: row['Shift'] || '-',
      tglT: row['Date'] || '-',
      wktT: row['Time'] || '-',
      tglC: row['Action Date'] || '-',
      wktC: row['Action Time'] || '-',
      mo: row['MO'] || '-',
      item: row['Item Code'] || '-',
      msn: row['MC'] || '-',
      tm: row['Team'] || '-',
      opr: row['Initiator'] || '-',
      alert: row['Quality Alert Type'] || row['Category Quality of Alert'] || '-',
      desc: row['Event Description (Penjelasan masalah)'] || '-',
      act: row['Action'] || '-',
      actOp: row['Action by'] || '-'
    }));

    return NextResponse.json(formattedData);
  } catch (error) {
    return NextResponse.json({ error: 'Gagal membaca data Excel' }, { status: 500 });
  }
}

export async function POST(request) {
  try {
    const newData = await request.json();
    let workbook;
    let data = [];
    let sheetName = 'Sheet1';

    // Buka file yang ada, atau buat baru kalau belum ada
    if (fs.existsSync(filePath)) {
      const fileBuffer = fs.readFileSync(filePath);
      workbook = XLSX.read(fileBuffer, { type: 'buffer' });
      sheetName = workbook.SheetNames[0];
      data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
    } else {
      workbook = XLSX.utils.book_new();
    }

    // Mapping dari form Dashboard untuk ditulis kembali ke Excel
    const newExcelRow = {
      'Date': newData.tglT,
      'Shift': newData.shift,
      'Time': newData.wktT,
      'MC': newData.msn,
      'Team': newData.tm,
      'Item Code': newData.item,
      'MO': newData.mo,
      'Alert Type': newData.alertType,
      'Quality Alert Type': newData.alert,
      'Event Description (Penjelasan masalah)': newData.desc,
      'Initiator': newData.opr,
      'Action Date': newData.tglC,
      'Action Time': newData.wktC,
      'Action': newData.act,
      'Action by': newData.actOp
    };

    // Tambahkan data ke baris paling bawah
    data.push(newExcelRow);
    
    // Tulis dan simpan perubahan ke file Excel
    const newSheet = XLSX.utils.json_to_sheet(data);
    workbook.Sheets[sheetName] = newSheet;
    XLSX.writeFile(workbook, filePath);

    return NextResponse.json({ success: true });
  } catch (error) {
    return NextResponse.json({ error: 'Gagal menyimpan ke Excel' }, { status: 500 });
  }
}