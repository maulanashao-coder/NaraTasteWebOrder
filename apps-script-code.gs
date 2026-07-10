function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    var name = e.parameter.name || '';
    var phone = e.parameter.phone || '';
    var address = e.parameter.address || '';
    var deliveryDate = e.parameter.delivery_date || '';
    var notes = e.parameter.notes || '';
    var orderSummary = e.parameter.order_summary || '';
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

    sheet.appendRow([new Date(), name, phone, address, deliveryDate, orderSummary, shippingCost, totalPrice, notes, photoUrl]);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCreateFolder(name) {
  var folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(name);
}
