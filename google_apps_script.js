/**
 * GOOGLE APPS SCRIPT FOR ATUR UANG DATABASE
 * 
 * Petunjuk Instalasi:
 * 1. Buka Google Sheets baru di Google Drive Anda.
 * 2. Klik Ekstensi -> Apps Script.
 * 3. Hapus semua kode bawaan dan tempel kode ini di bawah ini.
 * 4. Klik Simpan (ikon disket).
 * 5. Klik Terapkan -> Penerapan baru (Deploy -> New deployment).
 * 6. Pilih Jenis: Aplikasi Web (Web App).
 * 7. Isi Deskripsi penerapan, contoh: "Atur Uang API".
 * 8. Konfigurasi:
 *    - Jalankan sebagai: Saya (Email Anda)
 *    - Siapa yang memiliki akses: Siapa saja (Anyone) -> Ini penting agar aplikasi web front-end dapat mengaksesnya.
 * 9. Klik Terapkan (Deploy). Izinkan akses akun Google Anda jika diminta.
 * 10. Salin "URL Aplikasi Web" yang dihasilkan dan tempel di menu Pengaturan aplikasi Atur Uang Anda.
 */

// Konstanta Token Keamanan Opsional - Untuk melindungi database Anda dari penulisan tidak sah
var SECURITY_TOKEN = ""; // Kosongkan jika tidak ingin menggunakan token keamanan, atau isi dengan kata sandi acak

function doGet(e) {
    return handleResponse(function () {
        var key = e.parameter.key || "";
        if (SECURITY_TOKEN && key !== SECURITY_TOKEN) {
            return { success: false, error: "Akses Ditolak: Token keamanan tidak valid." };
        }

        var db = initDb();

        var collaborators = ["panji.newbie@gmail.com"];
        try {
            var ss = SpreadsheetApp.getActiveSpreadsheet();
            ss.getEditors().forEach(function (u) {
                var email = u.getEmail();
                if (email && collaborators.indexOf(email) === -1) {
                    collaborators.push(email);
                }
            });
            ss.getViewers().forEach(function (u) {
                var email = u.getEmail();
                if (email && collaborators.indexOf(email) === -1) {
                    collaborators.push(email);
                }
            });
        } catch (err) {
            // Mengabaikan jika izin Apps Script dibatasi
        }

        var data = {
            success: true,
            transactions: getSheetData(db.transactionsSheet),
            budgets: getSheetData(db.budgetsSheet),
            settings: getSheetData(db.settingsSheet),
            collaborators: collaborators
        };

        return data;
    });
}

function doPost(e) {
    return handleResponse(function () {
        var payload;
        try {
            payload = JSON.parse(e.postData.contents);
        } catch (err) {
            return { success: false, error: "Format JSON tidak valid." };
        }

        var key = payload.key || e.parameter.key || "";
        if (SECURITY_TOKEN && key !== SECURITY_TOKEN) {
            return { success: false, error: "Akses Ditolak: Token keamanan tidak valid." };
        }

        var action = payload.action;
        var db = initDb();

        switch (action) {
            case "syncAll":
                // Menimpa data lokal ke spreadsheet
                setSheetData(db.transactionsSheet, payload.transactions);
                setSheetData(db.budgetsSheet, payload.budgets);
                setSheetData(db.settingsSheet, payload.settings);
                return { success: true, message: "Sinkronisasi penuh berhasil." };

            case "addTransaction":
                appendRow(db.transactionsSheet, payload.data);
                return { success: true, message: "Transaksi berhasil ditambahkan." };

            case "updateTransaction":
                updateRow(db.transactionsSheet, "ID", payload.data.ID, payload.data);
                return { success: true, message: "Transaksi berhasil diperbarui." };

            case "deleteTransaction":
                deleteRow(db.transactionsSheet, "ID", payload.id);
                return { success: true, message: "Transaksi berhasil dihapus." };

            case "saveBudget":
                upsertRow(db.budgetsSheet, "Kategori", payload.data.Kategori, payload.data);
                return { success: true, message: "Anggaran berhasil disimpan." };

            case "deleteBudget":
                deleteRow(db.budgetsSheet, "Kategori", payload.kategori);
                return { success: true, message: "Anggaran berhasil dihapus." };

            case "saveSetting":
                upsertRow(db.settingsSheet, "Key", payload.data.Key, payload.data);
                return { success: true, message: "Pengaturan berhasil disimpan." };

            default:
                return { success: false, error: "Aksi tidak dikenal: " + action };
        }
    });
}

// ================= API HELPERS =================

function handleResponse(callback) {
    var output;
    try {
        var result = callback();
        output = JSON.stringify(result);
    } catch (err) {
        output = JSON.stringify({ success: false, error: err.toString() });
    }

    return ContentService.createTextOutput(output)
        .setMimeType(ContentService.MimeType.JSON);
}

function initDb() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var tSheet = ss.getSheetByName("Transactions");
    if (!tSheet) {
        tSheet = ss.insertSheet("Transactions");
        var headers = ["ID", "Tanggal", "Kategori", "Tipe", "Keterangan", "Jumlah", "Arus"];
        tSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
        tSheet.setFrozenRows(1);
    }

    var bSheet = ss.getSheetByName("Budgets");
    if (!bSheet) {
        bSheet = ss.insertSheet("Budgets");
        var headers = ["Kategori", "Limit", "Tipe"];
        bSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
        bSheet.setFrozenRows(1);
    }

    var sSheet = ss.getSheetByName("Settings");
    if (!sSheet) {
        sSheet = ss.insertSheet("Settings");
        var headers = ["Key", "Value"];
        sSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
        sSheet.setFrozenRows(1);
    }

    // Hapus Sheet default jika ada dan kosong
    var defaultSheet = ss.getSheetByName("Sheet1");
    if (defaultSheet && defaultSheet.getLastRow() === 0) {
        ss.deleteSheet(defaultSheet);
    }

    return {
        transactionsSheet: tSheet,
        budgetsSheet: bSheet,
        settingsSheet: sSheet
    };
}

function getSheetData(sheet) {
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

    var data = [];
    for (var r = 0; r < values.length; r++) {
        var obj = {};
        for (var c = 0; c < headers.length; c++) {
            var val = values[r][c];
            // Konversi tipe data agar sesuai
            if (headers[c] === "Jumlah" || headers[c] === "Limit") {
                obj[headers[c]] = Number(val) || 0;
            } else if (val instanceof Date) {
                obj[headers[c]] = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
            } else {
                obj[headers[c]] = val;
            }
        }
        data.push(obj);
    }
    return data;
}

function setSheetData(sheet, dataList) {
    var lastRow = sheet.getLastRow();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Bersihkan data lama
    if (lastRow > 1) {
        sheet.deleteRows(2, lastRow - 1);
    }

    if (!dataList || dataList.length === 0) return;

    var values = [];
    for (var i = 0; i < dataList.length; i++) {
        var rowObj = dataList[i];
        var rowVals = [];
        for (var c = 0; c < headers.length; c++) {
            var val = rowObj[headers[c]];
            rowVals.push(val === undefined ? "" : val);
        }
        values.push(rowVals);
    }

    sheet.getRange(2, 1, values.length, headers.length).setValues(values);
}

function appendRow(sheet, rowData) {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var rowVals = [];
    for (var c = 0; c < headers.length; c++) {
        var val = rowData[headers[c]];
        rowVals.push(val === undefined ? "" : val);
    }
    sheet.appendRow(rowVals);
}

function updateRow(sheet, primaryKeyName, primaryKeyValue, rowData) {
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var pkColIndex = headers.indexOf(primaryKeyName) + 1;
    if (pkColIndex === 0) return;

    var values = sheet.getRange(2, pkColIndex, lastRow - 1, 1).getValues();
    for (var r = 0; r < values.length; r++) {
        if (values[r][0].toString() === primaryKeyValue.toString()) {
            var rowIndex = r + 2;
            var rowVals = [];
            for (var c = 0; c < headers.length; c++) {
                var val = rowData[headers[c]];
                rowVals.push(val === undefined ? "" : val);
            }
            sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowVals]);
            break;
        }
    }
}

function deleteRow(sheet, primaryKeyName, primaryKeyValue) {
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var pkColIndex = headers.indexOf(primaryKeyName) + 1;
    if (pkColIndex === 0) return;

    var values = sheet.getRange(2, pkColIndex, lastRow - 1, 1).getValues();
    for (var r = values.length - 1; r >= 0; r--) {
        if (values[r][0].toString() === primaryKeyValue.toString()) {
            sheet.deleteRow(r + 2);
        }
    }
}

function upsertRow(sheet, primaryKeyName, primaryKeyValue, rowData) {
    var lastRow = sheet.getLastRow();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var pkColIndex = headers.indexOf(primaryKeyName) + 1;
    if (pkColIndex === 0) return;

    var found = false;
    if (lastRow > 1) {
        var values = sheet.getRange(2, pkColIndex, lastRow - 1, 1).getValues();
        for (var r = 0; r < values.length; r++) {
            if (values[r][0].toString() === primaryKeyValue.toString()) {
                var rowIndex = r + 2;
                var rowVals = [];
                for (var c = 0; c < headers.length; c++) {
                    var val = rowData[headers[c]];
                    rowVals.push(val === undefined ? "" : val);
                }
                sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowVals]);
                found = true;
                break;
            }
        }
    }

    if (!found) {
        appendRow(sheet, rowData);
    }
}
