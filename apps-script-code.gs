// Fixed info columns that always come first (A through J)
var FIXED_COLUMNS = ['Waktu', 'Nama', 'Telepon', 'Alamat', 'Tanggal Pengiriman', 'Total Harga Pesanan', 'Ongkos Kirim', 'Total Harga', 'Catatan', 'Link Foto'];

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    var name = e.parameter.name || '';
    var phone = e.parameter.phone || '';
    var address = e.parameter.address || '';
    var deliveryDate = e.parameter.delivery_date || '';
    var notes = e.parameter.notes || '';
    var orderSummary = e.parameter.order_summary || '';
    var itemsSubtotal = e.parameter.items_subtotal || '';
    var shippingCost = e.parameter.shipping_cost_display || '';
    var totalPrice = e.parameter.total_price || '';

    var photoUrl = '';
    var photoBase64 = e.parameter.photo_base64 || '';
    if (photoBase64) {
      var filename = e.parameter.photo_filename || 'photo.jpg';
      var mimeType = e.parameter.photo_mimetype || 'image/jpeg';
      var folder = getOrCreateFolder('Foto Pesanan Nara Taste');
      var decodedBytes = Utilities.base64Decode(photoBase64);
      var blob = Utilities.newBlob(decodedBytes, mimeType, filename);
      var driveFile = folder.createFile(blob);
      driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      photoUrl = driveFile.getUrl();
    }

    // Make sure the fixed header row exists
    var lastRow = sheet.getLastRow();
    if (lastRow === 0) {
      sheet.getRange(1, 1, 1, FIXED_COLUMNS.length).setValues([FIXED_COLUMNS]);
      lastRow = 1;
    }

    var newRow = lastRow + 1;

    // Write the fixed info columns (A-J)
    sheet.getRange(newRow, 1, 1, FIXED_COLUMNS.length).setValues([[
      new Date(), name, phone, address, deliveryDate, itemsSubtotal, shippingCost, totalPrice, notes, photoUrl
    ]]);

    // Parse "Item Name xQty" lines and place quantity into that item's column
    console.log('orderSummary raw: ' + orderSummary);
    var lines = orderSummary.split('\n').filter(function (l) { return l.trim() !== ''; });
    console.log('lines parsed: ' + JSON.stringify(lines));
    lines.forEach(function (line) {
      var match = line.match(/^(.*) x(\d+)$/);
      console.log('line: "' + line + '" match: ' + JSON.stringify(match));
      if (match) {
        var itemName = match[1].trim();
        var qty = parseInt(match[2], 10);
        var col = getOrCreateItemColumn(sheet, itemName);
        console.log('itemName: "' + itemName + '" qty: ' + qty + ' col: ' + col);
        sheet.getRange(newRow, col).setValue(qty);
      }
    });

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Finds the column for a given menu item by matching the header row (row 1).
// If the item doesn't have a column yet, creates a new one at the end.
function getOrCreateItemColumn(sheet, itemName) {
  var lastCol = Math.max(sheet.getLastColumn(), FIXED_COLUMNS.length);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  for (var i = FIXED_COLUMNS.length; i < headers.length; i++) {
    if (headers[i] === itemName) {
      return i + 1; // convert to 1-indexed column number
    }
  }

  var newCol = lastCol + 1;
  sheet.getRange(1, newCol).setValue(itemName);
  return newCol;
}

function getOrCreateFolder(name) {
  var folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(name);
}

// TEMPORARY TEST FUNCTION — run this directly in the editor to debug.
// Select "testDoPost" from the function dropdown at the top, then click Run.
// Check the log output panel that appears immediately at the bottom of the screen.
function testDoPost() {
  var mockE = {
    parameter: {
      name: 'Test User',
      phone: '0812xxxx',
      address: 'Jl Testing No 1',
      delivery_date: '2026-07-20',
      notes: 'ini test',
      order_summary: 'Mango Sticky Rice x2\nEs Pisang Ijo x1',
      items_subtotal: 'Rp 65.000',
      shipping_cost_display: 'Rp 10.000',
      total_price: 'Rp 75.000'
    }
  };
  var result = doPost(mockE);
  console.log('RESULT: ' + result.getContent());
}
