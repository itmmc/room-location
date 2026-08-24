/* =================================================================
   Modern Mills - Room Location Add-in
   -----------------------------------------------------------------
   يضيف عنوان الشركة + رابط الخريطة + اسم المنظّم + رقم الاستقبال
   الشرط: فقط إذا كان الاجتماع فيه غرفة محجوزة
   ================================================================= */


/* ============ الإعدادات - عدّل هنا فقط ============ */

var CONFIG = {

  mapUrl: "https://maps.google.com/?q=21.548840,39.136276",

  addressLine1: "Modern Mills Company &ndash; HQ",
  addressLine2: "Moh. Attaib Attunsi St, Jeddah, Makkah JEKD2698, SA",

  // رقم الاستقبال - اتركه فاضي "" لإخفاء السطر
  // لو حبيت تفعّله لاحقاً، حط الرقم هنا وارفع الملف من جديد
  receptionPhone: "",
  receptionLabel: "Reception",

  // الغرف المفعّلة. القائمة الفاضية [] تعني كل الغرف.
  // للتجربة: غرفة وحدة فقط
  onlyRooms: [
    "intelligence@modernmills.com.sa"
  ]
};

/* ================================================== */


var MARKER = "MM-LOC-BLOCK";


Office.onReady(function () { });


/* ---------- بناء البلوك ---------- */

function buildBlock(organizerName) {

  var rows = [];

  if (organizerName) {
    rows.push("Organizer: " + organizerName);
  }

  if (CONFIG.receptionPhone) {
    rows.push(
      CONFIG.receptionLabel + ': <a href="tel:' +
      CONFIG.receptionPhone.replace(/\s/g, "") +
      '" style="color:#185FA5;">' + CONFIG.receptionPhone + "</a>"
    );
  }

  var footer = rows.length
    ? '<div style="margin-top:8px;padding-top:6px;border-top:1px solid #B5D4F4;' +
      'font-size:12px;color:#185FA5;line-height:1.6;">' +
      rows.join("<br>") + "</div>"
    : "";

  return '' +
  '<div id="' + MARKER + '" style="margin-top:16px;padding:12px 14px;' +
  'border:1px solid #B5D4F4;border-radius:6px;background:#E6F1FB;' +
  'font-family:Segoe UI,Arial,sans-serif;">' +
    '<div style="font-size:13px;font-weight:600;color:#0C447C;margin-bottom:5px;">' +
      'Meeting location / &#1605;&#1608;&#1602;&#1593; &#1575;&#1604;&#1575;&#1580;&#1578;&#1605;&#1575;&#1593;' +
    '</div>' +
    '<div style="font-size:12px;color:#185FA5;line-height:1.6;">' +
      CONFIG.addressLine1 + '<br>' + CONFIG.addressLine2 +
    '</div>' +
    '<div style="margin-top:6px;">' +
      '<a href="' + CONFIG.mapUrl + '" style="font-size:12px;color:#185FA5;">' +
        'Open in Google Maps' +
      '</a>' +
    '</div>' +
    footer +
  '</div>';
}


/* ---------- نقطة الدخول ---------- */

function onAppointmentSendHandler(event) {

  var item = Office.context.mailbox.item;

  if (!item || !item.enhancedLocation) {
    event.completed({ allowEvent: true });
    return;
  }

  item.enhancedLocation.getAsync(function (locResult) {

    if (locResult.status !== Office.AsyncResultStatus.Succeeded) {
      event.completed({ allowEvent: true });
      return;
    }

    var hasRoom = false;
    var locations = locResult.value || [];
    var filter = CONFIG.onlyRooms || [];

    for (var i = 0; i < locations.length; i++) {

      if (locations[i].type !== Office.MailboxEnums.LocationType.Room) {
        continue;
      }

      // القائمة فاضية -> كل الغرف مفعّلة
      if (filter.length === 0) {
        hasRoom = true;
        break;
      }

      var addr = (locations[i].emailAddress || "").toLowerCase();

      for (var j = 0; j < filter.length; j++) {
        if (filter[j].toLowerCase() === addr) {
          hasRoom = true;
          break;
        }
      }

      if (hasRoom) { break; }
    }

    // ما فيه غرفة مفعّلة -> لا تتدخل إطلاقاً
    if (!hasRoom) {
      event.completed({ allowEvent: true });
      return;
    }

    appendBlock(item, event);
  });
}


/* ---------- إدراج البلوك ---------- */

function appendBlock(item, event) {

  var name = "";
  try {
    name = Office.context.mailbox.userProfile.displayName || "";
  } catch (e) { }

  item.body.getAsync(Office.CoercionType.Html, function (bodyResult) {

    if (bodyResult.status === Office.AsyncResultStatus.Succeeded &&
        bodyResult.value &&
        bodyResult.value.indexOf(MARKER) !== -1) {
      event.completed({ allowEvent: true });
      return;
    }

    item.body.appendOnSendAsync(
      buildBlock(name),
      { coercionType: Office.CoercionType.Html },
      function () {
        event.completed({ allowEvent: true });
      }
    );
  });
}


Office.actions.associate("onAppointmentSendHandler", onAppointmentSendHandler);
