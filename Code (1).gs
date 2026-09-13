function doGet(e) {
  var action = e.parameter.action;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (action === "getItems") {
    var sheet = ss.getSheetByName("Items");
    var rows = sheet.getDataRange().getValues();
    var items = [];
    // افترض أن الأعمدة: Code, Name, Type, Price, Stock
    for (var i = 1; i < rows.length; i++) {
      if(rows[i][0] !== "" && rows[i][0] !== null && rows[i][0] !== undefined) {
        items.push({
          code: rows[i][0].toString().trim(),
          name: rows[i][1],
          type: rows[i][2], // 'piece' أو 'weight'
          price: rows[i][3],
          stock: rows[i][4]
        });
      }
    }
    return ContentService.createTextOutput(JSON.stringify(items)).setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action === "getInvoices") {
    var sheet = ss.getSheetByName("Invoices");
    var rows = sheet.getDataRange().getValues();
    var invoices = [];
    for (var i = 1; i < rows.length; i++) {
      if(rows[i][0] !== "" && rows[i][0] !== null && rows[i][0] !== undefined) {
        invoices.push({
          invoiceID: rows[i][0].toString().trim(),
          date: rows[i][1],
          time: rows[i][2],
          customerName: rows[i][3],
          customerPhone: rows[i][4],
          subTotal: rows[i][5],
          discountType: rows[i][6],
          discountValue: rows[i][7],
          netTotal: rows[i][8],
          paymentDetails: rows[i][9],
          status: rows[i][10],
          cashierName: rows[i][11]
        });
      }
    }
    return ContentService.createTextOutput(JSON.stringify(invoices)).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = JSON.parse(e.postData.contents);
  var action = data.action;

  // ================== تسجيل الدخول (كان ناقص، وهو سبب تعطل الدخول بالكامل) ==================
  if (action === "login") {
    var sheet = ss.getSheetByName("Users");

    // لو شيت المستخدمين مش موجود أصلاً
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        message: "لا يوجد شيت باسم Users في ملف الإكسيل"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    var rows = sheet.getDataRange().getValues();
    // افترض أن الأعمدة: Username, Password, Role
    var inputUser = data.username !== null && data.username !== undefined ? data.username.toString().trim() : "";
    var inputPass = data.password !== null && data.password !== undefined ? data.password.toString().trim() : "";

    var found = null;
    for (var i = 1; i < rows.length; i++) {
      var rowUser = rows[i][0] !== null && rows[i][0] !== undefined ? rows[i][0].toString().trim() : "";
      var rowPass = rows[i][1] !== null && rows[i][1] !== undefined ? rows[i][1].toString().trim() : "";

      if (rowUser !== "" && rowUser === inputUser && rowPass === inputPass) {
        found = { username: rowUser, role: rows[i][2] !== null && rows[i][2] !== undefined ? rows[i][2].toString().trim() : "" };
        break;
      }
    }

    if (found) {
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        username: found.username,
        role: found.role
      })).setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        message: "اسم المستخدم أو كلمة المرور غير صحيحة"
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  // ================================================================================
  
  if (action === "saveInvoice") {
    var sheet = ss.getSheetByName("Invoices");
    sheet.appendRow([
      data.invoiceID,
      data.date,
      data.time,
      data.customerName,
      data.customerPhone,
      data.subTotal,
      data.discountType,
      data.discountValue,
      data.netTotal,
      data.paymentDetails,
      data.status,
      data.cashierName
    ]);
    
    // حفظ تفاصيل الأصناف في InvoiceDetails وتخصيم المخزون
    var detailsSheet = ss.getSheetByName("InvoiceDetails");
    var itemsSheet = ss.getSheetByName("Items");
    var itemsRows = itemsSheet.getDataRange().getValues();
    
    data.items.forEach(function(item) {
      detailsSheet.appendRow([
        data.invoiceID,
        item.code,
        item.name,
        item.type,
        item.qty,
        item.displayQty,
        item.price,
        item.price * item.qty
      ]);
      
      // تحديث المخزون في شيت Items
      for (var i = 1; i < itemsRows.length; i++) {
        var sheetCode = itemsRows[i][0] !== null && itemsRows[i][0] !== undefined ? itemsRows[i][0].toString().trim() : "";
        var targetCode = item.code !== null && item.code !== undefined ? item.code.toString().trim() : "";
        
        if (sheetCode === targetCode) {
          var currentStock = parseFloat(itemsRows[i][4]) || 0;
          var newStock = currentStock - item.qty;
          itemsSheet.getRange(i + 1, 5).setValue(newStock < 0 ? 0 : newStock);
          break;
        }
      }
    });
    
    return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action === "updateInvoiceStatus") {
    var sheet = ss.getSheetByName("Invoices");
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      var rowInvoiceID = rows[i][0] !== null && rows[i][0] !== undefined ? rows[i][0].toString().trim() : "";
      var targetInvoiceID = data.invoiceID !== null && data.invoiceID !== undefined ? data.invoiceID.toString().trim() : "";
      
      if (rowInvoiceID === targetInvoiceID) {
        sheet.getRange(i + 1, 11).setValue(data.status); // عمود Status
        break;
      }
    }
    return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action === "saveVault") {
    var sheet = ss.getSheetByName("Vault_Log");
    sheet.appendRow([
      data.date,
      data.time,
      data.type, // مصاريف / إيرادات
      data.amount,
      data.reason,
      data.cashier
    ]);
    return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
  }
}
