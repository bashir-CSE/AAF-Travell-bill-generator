var SPREADSHEET_ID = "1tPnn7Du2Bt99HNTe86N8B0KBTqAA0KfZl69ycpf2ZPg";
/* Parent folder: "Automation Generated File" — ID remains unchanged */
var DRIVE_FOLDER_ID = "1DvAqAkA5FzpCFXOpyO6_0mLEuuJvJrEY";
var SHEET_SECTIONS = "Sections list";
var SHEET_ARTISANS = "Artisans list";
var SHEET_BILLS = "History data";

/* ── মডিউল-লেভেল স্প্রেডশিট ক্যাশ ── */
var _cachedSpreadsheet = null;

function getCachedSpreadsheet() {
	if (!_cachedSpreadsheet) {
		_cachedSpreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
	}
	return _cachedSpreadsheet;
}

function doGet() {
	return HtmlService.createTemplateFromFile("index")
		.evaluate()
		.setTitle("যাতায়াত বিল")
		.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function ensureCleanupTrigger() {
	var triggers = ScriptApp.getProjectTriggers();
	var exists = false;
	for (var i = 0; i < triggers.length; i++) {
		if (triggers[i].getHandlerFunction() === "midnightCleanup") {
			exists = true;
			break;
		}
	}
	if (!exists) {
		ScriptApp.newTrigger("midnightCleanup")
			.timeBased()
			.atHour(0)
			.everyDays(1)
			.create();
	}
}

/* ══════════════════════════════════════════════════════════════════════
   midnightCleanup — runs daily at 00:00
   History sheet structure: Col A = Section Name, Col B = Bill URL
   No date/time columns — just trash all files and clear all rows.
   ══════════════════════════════════════════════════════════════════════ */
function midnightCleanup() {
	try {
		var ss = getCachedSpreadsheet();
		var sheet = ss.getSheetByName(SHEET_BILLS);
		if (!sheet) return;
		if (sheet.getLastRow() <= 1) return;

		var lastRow = sheet.getLastRow();
		var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

		for (var i = 0; i < data.length; i++) {
			var fileUrl = String(data[i][1] || "");
			if (fileUrl === "" || fileUrl === "undefined") continue;
			try {
				var match = fileUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
				if (match) {
					DriveApp.getFileById(match[1]).setTrashed(true);
				}
			} catch (e) {
				/* File may already be trashed or access lost — skip silently */
			}
		}

		SpreadsheetApp.flush();

		for (var j = lastRow; j >= 2; j--) {
			try {
				sheet.deleteRow(j);
			} catch (e) {}
		}

		/* ক্যাশ মুছে ফেলুন যাতে পরবর্তী রিকোয়েস্টে সতেজ ডেটা আসে */
		clearHistoryCache();
	} catch (e) {
		Logger.log("midnightCleanup error: " + e.toString());
	}
}

function toBanglaNum(input) {
	var banglaDigits = ["০", "১", "২", "৩", "৪", "৫", "৬", "৭", "৮", "৯"];
	return String(input).replace(/[0-9]/g, function (d) {
		return banglaDigits[parseInt(d, 10)];
	});
}

function numberToBanglaWords(num) {
	if (num === 0) return "শূন্য";
	if (num < 0) return "ঋণাত্মক " + numberToBanglaWords(Math.abs(num));

	var ones = [
		"",
		"এক",
		"দুই",
		"তিন",
		"চার",
		"পাঁচ",
		"ছয়",
		"সাত",
		"আট",
		"নয়",
		"দশ",
		"এগারো",
		"বারো",
		"তেরো",
		"চৌদ্দ",
		"পনেরো",
		"ষোল",
		"সতেরো",
		"আঠারো",
		"উনিশ",
		"বিশ",
		"একুশ",
		"বাইশ",
		"তেইশ",
		"চব্বিশ",
		"পঁচিশ",
		"ছাব্বিশ",
		"সাতাশ",
		"আটাশ",
		"উনত্রিশ",
		"ত্রিশ",
		"একত্রিশ",
		"বত্রিশ",
		"তেত্রিশ",
		"চৌত্রিশ",
		"পঁয়ত্রিশ",
		"ছত্রিশ",
		"সাতত্রিশ",
		"আটত্রিশ",
		"উনচল্লিশ",
		"চল্লিশ",
		"একচল্লিশ",
		"বিয়াল্লিশ",
		"তেতাল্লিশ",
		"চউচল্লিশ",
		"পয়তাল্লিশ",
		"ছেচল্লিশ",
		"সাতচল্লিশ",
		"আটচল্লিশ",
		"উনপঞ্চাশ",
		"পঞ্চাশ",
		"একান্ন",
		"বায়ান্ন",
		"তিপান্ন",
		"চুয়ান্ন",
		"পঞ্চান্ন",
		"ছাপান্ন",
		"সাতান্ন",
		"আটান্ন",
		"উনষাট",
		"ষাট",
		"একষট্টি",
		"বাষট্টি",
		"তেষট্টি",
		"চৌষট্টি",
		"পয়ষট্টি",
		"ছেষট্টি",
		"সাতষট্টি",
		"আটষট্টি",
		"উনসত্তর",
		"সত্তর",
		"একাত্তর",
		"বাহাত্তর",
		"তেহাত্তর",
		"চুরাত্তর",
		"পচাত্তর",
		"ছিয়াত্তর",
		"সাতাত্তর",
		"আটাত্তর",
		"উনাশি",
		"আশি",
		"একাশি",
		"বিরাশি",
		"তিরাশি",
		"চুরাশি",
		"পঁচাশি",
		"ছিয়াশি",
		"সাতাশি",
		"আটাশি",
		"উননব্বই",
		"নব্বই",
		"একানব্বই",
		"বিরানব্বই",
		"তিরানব্বই",
		"চুরানব্বই",
		"পঁচানব্বই",
		"ছিয়ানব্বই",
		"সাতানব্বই",
		"আটানব্বই",
		"নিরানব্বই",
	];

	if (num < 100) return ones[num];

	var result = [];
	var n = Math.floor(num);

	if (n >= 10000000) {
		result.push(numberToBanglaWords(Math.floor(n / 10000000)) + " কোটি");
		n %= 10000000;
	}
	if (n >= 100000) {
		result.push(numberToBanglaWords(Math.floor(n / 100000)) + " লাখ");
		n %= 100000;
	}
	if (n >= 1000) {
		result.push(numberToBanglaWords(Math.floor(n / 1000)) + " হাজার");
		n %= 1000;
	}
	if (n >= 100) {
		result.push(ones[Math.floor(n / 100)] + " শত");
		n %= 100;
	}
	if (n > 0) {
		result.push(ones[n]);
	}

	return result.join(" ");
}

/* ── Helper: "2025-06" → Bengali month name ── */
function getBanglaMonthName(monthStr) {
	var months = [
		"জানুয়ারি",
		"ফেব্রুয়ারি",
		"মার্চ",
		"এপ্রিল",
		"মে",
		"জুন",
		"জুলাই",
		"আগস্ট",
		"সেপ্টেম্বর",
		"অক্টোবর",
		"নভেম্বর",
		"ডিসেম্বর",
	];
	var parts = monthStr.split("-");
	var monthIndex = parseInt(parts[1], 10) - 1;
	return months[monthIndex];
}

function getSections() {
	var ss = getCachedSpreadsheet();
	var sheet = ss.getSheetByName(SHEET_SECTIONS);
	if (!sheet) return [];
	var lastRow = sheet.getLastRow();
	if (lastRow < 2) return [];
	var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
	var sections = [];
	for (var i = 0; i < data.length; i++) {
		if (data[i][0]) sections.push(data[i][0]);
	}
	return sections;
}

function getInitialData() {
	ensureCleanupTrigger();
	var sections = getSections();
	/* travelRate আর সেভ করা হচ্ছে না — প্রতিবার ব্যবহারকারীকে দিতে হবে */
	return { sections: sections };
}

/* ══════════════════════════════════════════════════════════════════════
   হিস্ট্রি ক্যাশিং — CacheService ব্যবহার করে ২০ সেকেন্ডের জন্য ক্যাশ রাখা হয়
   ══════════════════════════════════════════════════════════════════════ */
function clearHistoryCache() {
	CacheService.getScriptCache().remove("history_data");
}

function getHistory() {
	/* প্রথমে ক্যাশ চেক করুন */
	var cache = CacheService.getScriptCache();
	var cached = cache.get("history_data");
	if (cached) {
		return JSON.parse(cached);
	}

	var ss = getCachedSpreadsheet();
	var sheet = ss.getSheetByName(SHEET_BILLS);
	if (!sheet || sheet.getLastRow() <= 1) return [];

	var lastRow = sheet.getLastRow();
	var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

	var bills = [];
	for (var i = data.length - 1; i >= 0; i--) {
		var section = String(data[i][0] || "").trim();
		var url = String(data[i][1] || "").trim();
		if (section && url) {
			bills.push({
				section: section,
				url: url,
			});
		}
	}

	/* ২০ সেকেন্ডের জন্য ক্যাশ সেভ করুন */
	cache.put("history_data", JSON.stringify(bills), 20);
	return bills;
}

function getArtisans(sectionName) {
	var ss = getCachedSpreadsheet();
	var sheet = ss.getSheetByName(SHEET_ARTISANS);
	if (!sheet) return [];
	var lastRow = sheet.getLastRow();
	if (lastRow < 2) return [];
	var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
	var artisans = [];
	for (var i = 0; i < data.length; i++) {
		if (String(data[i][0]).trim() === sectionName.trim() && data[i][1]) {
			artisans.push({ name: data[i][1], code: data[i][2] || "" });
		}
	}
	return artisans;
}

function addArtisan(sectionName, artisanName, artisanCode) {
	if (!sectionName || !artisanName) {
		return { success: false, message: "সেকশনের নাম এবং কর্মীর নাম আবশ্যক।" };
	}

	var ss = getCachedSpreadsheet();
	var sheet = ss.getSheetByName(SHEET_ARTISANS);
	if (!sheet) {
		return {
			success: false,
			message: "স্প্রেডশিটে কর্মী তালিকা পাওয়া যায়নি।",
		};
	}

	var lastRow = sheet.getLastRow();
	if (lastRow >= 2) {
		var existingData = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
		for (var i = 0; i < existingData.length; i++) {
			if (
				String(existingData[i][0]).trim().toLowerCase() ===
					sectionName.trim().toLowerCase() &&
				String(existingData[i][1]).trim().toLowerCase() ===
					artisanName.trim().toLowerCase()
			) {
				return {
					success: false,
					message: '"' + artisanName + '" ইতিমধ্যে এই সেকশনে আছে।',
				};
			}
		}
	}

	sheet.appendRow([
		sectionName.trim(),
		artisanName.trim(),
		artisanCode ? artisanCode.trim() : "",
	]);
	var updatedArtisans = getArtisans(sectionName);
	return {
		success: true,
		message: "কর্মী সফলভাবে যোগ হয়েছে!",
		artisans: updatedArtisans,
	};
}

function removeArtisan(sectionName, artisanName) {
	if (!sectionName || !artisanName) {
		return { success: false, message: "সেকশনের নাম এবং কর্মীর নাম আবশ্যক।" };
	}

	var ss = getCachedSpreadsheet();
	var sheet = ss.getSheetByName(SHEET_ARTISANS);
	if (!sheet) {
		return {
			success: false,
			message: "স্প্রেডশিটে কর্মী তালিকা পাওয়া যায়নি।",
		};
	}

	var lastRow = sheet.getLastRow();
	if (lastRow < 2) {
		return { success: false, message: "কোনো কর্মী পাওয়া যায়নি।" };
	}

	var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
	var rowToDelete = -1;
	for (var i = 0; i < data.length; i++) {
		if (
			String(data[i][0]).trim().toLowerCase() ===
				sectionName.trim().toLowerCase() &&
			String(data[i][1]).trim().toLowerCase() ===
				artisanName.trim().toLowerCase()
		) {
			rowToDelete = i + 2;
			break;
		}
	}

	if (rowToDelete === -1) {
		return {
			success: false,
			message: '"' + artisanName + '" এই সেকশনে পাওয়া যায়নি।',
		};
	}

	sheet.deleteRow(rowToDelete);
	var updatedArtisans = getArtisans(sectionName);
	return {
		success: true,
		message: "কর্মী সফলভাবে সরানো হয়েছে!",
		artisans: updatedArtisans,
	};
}

function deleteBill(billUrl) {
	var ss = getCachedSpreadsheet();
	var sheet = ss.getSheetByName(SHEET_BILLS);
	if (!sheet || sheet.getLastRow() <= 1) {
		return { success: false, message: "মুছে ফেলার মতো কোনো বিল নেই।" };
	}

	var lastRow = sheet.getLastRow();
	var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

	var rowToDelete = -1;
	for (var i = 0; i < data.length; i++) {
		if (String(data[i][1]) === billUrl) {
			rowToDelete = i + 2;
			break;
		}
	}

	if (rowToDelete === -1) {
		return { success: false, message: "রেকর্ডে বিল পাওয়া যায়নি।" };
	}

	try {
		var match = billUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
		if (match) {
			DriveApp.getFileById(match[1]).setTrashed(true);
		}
	} catch (e) {}

	sheet.deleteRow(rowToDelete);

	/* হিস্ট্রি ক্যাশ মুছে ফেলুন যাতে পরবর্তী রিফ্রেশে আপডেটেড ডেটা আসে */
	clearHistoryCache();

	return { success: true, message: "বিল সফলভাবে মুছে ফেলা হয়েছে।" };
}

/* ══════════════════════════════════════════════════════════════════════
   generateBills
   History sheet: only 2 columns — Section Name, Bill URL
   No date, no time, no setNumberFormat — completely clean.
   ══════════════════════════════════════════════════════════════════════ */
function generateBills(selectedData, sectionName, monthStr, travelRate) {
	try {
		/* travelRate আর PropertiesService-এ সেভ হচ্ছে না */

		/* ── Find or create "Travel Bill" folder inside parent ── */
		var parentFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
		var targetFolder;
		var folders = parentFolder.getFoldersByName("Travel Bill");
		if (folders.hasNext()) {
			targetFolder = folders.next();
		} else {
			targetFolder = parentFolder.createFolder("Travel Bill");
		}

		/* ── Prepare display values ── */
		var now = new Date();
		var banglaMonthName = getBanglaMonthName(monthStr);
		var yearPart = toBanglaNum(monthStr.split("-")[0]);
		var banglaMonthDisplay = banglaMonthName + " " + yearPart;
		var timestamp = Utilities.formatDate(
			now,
			Session.getScriptTimeZone(),
			"dd/MM/yyyy, hh:mm a",
		);

		/* ── Generate PDF ── */
		var html = getPdfHtmlTemplate(
			selectedData,
			sectionName,
			banglaMonthDisplay,
			travelRate,
			timestamp,
		);
		var fileName =
			"Travel Bill - " +
			sectionName +
			" - " +
			banglaMonthName +
			" " +
			monthStr.split("-")[0] +
			".pdf";

		var blob = Utilities.newBlob(html, MimeType.HTML).setName(fileName);
		var pdfBlob = blob.getAs(MimeType.PDF);
		var file = targetFolder.createFile(pdfBlob);
		var fileUrl = file.getUrl();

		/* ── Write to History sheet — only 2 columns, no date/time ── */
		var ss = getCachedSpreadsheet();
		var historySheet = ss.getSheetByName(SHEET_BILLS);
		if (!historySheet) {
			historySheet = ss.insertSheet(SHEET_BILLS);
			historySheet.appendRow(["Section Name", "Bill"]);
			historySheet.getRange("A1:B1").setFontWeight("bold");
		}

		historySheet.appendRow([sectionName, fileUrl]);

		/* হিস্ট্রি ক্যাশ মুছে ফেলুন */
		clearHistoryCache();

		return {
			success: true,
			message: "আপনার যাতায়াত বিল সফলভাবে তৈরি হয়েছে",
			url: fileUrl,
		};
	} catch (error) {
		return { success: false, message: error.toString() };
	}
}

/* ══════════════════════════════════════════════════════════════════════
   getPdfHtmlTemplate
   ══════════════════════════════════════════════════════════════════════ */
function getPdfHtmlTemplate(
	data,
	sectionName,
	banglaMonthDisplay,
	travelRate,
	timestamp,
) {
	var banglaRate = toBanglaNum(travelRate);

	var html =
		"<html><head>" +
		'<link href="https://fonts.googleapis.com/css2?family=Tiro+Bangla&display=swap" rel="stylesheet">' +
		"<style>" +
		'@import url("https://fonts.googleapis.com/css2?family=Tiro+Bangla&display=swap");\n' +
		"@page { margin-bottom: 50px; }" +
		'body { font-family: "Tiro Bangla", "Inter", Arial, sans-serif; font-size: 9px; margin-top: 10; padding: 10px; padding-bottom: 90px; }' +
		"h1 { text-align: center; margin: 0; font-size: 16px; font-weight: bold; }" +
		"h2 { text-align: center; margin: 0; font-size: 14px; font-weight: bold; }" +
		"h3 { text-align: center; margin: 5px 0 10px 0; font-size: 12px; font-weight: bold; }" +
		"table { width: 100%; border-collapse: collapse; }" +
		"th, td { font-size: 12px; border: 0.5px solid #8b8b8b; padding: 1.3px; text-align: center; }" +
		"th { font-weight: bold; font-size: 12px;}" +
		".header-table { margin-bottom: -1px; }" +
		".header-table td { text-align: left; font-weight: bold; border-bottom: none; }" +
		".timestamp { text-align: right; font-size: 9px; color: #666; margin-bottom: 6px; }" +
		".footer-table { position: fixed; bottom: 0px; left: 10px; right: 10px; width: calc(100% - 10px); border: none; }" +
		".footer-table td { border: none; text-align: center; width: 25%; font-size: 10px; font-weight: bold; }" +
		".footer-table span { border-top: 1px dotted #000; padding-top: 3px; display: inline-block; width: 80%; }" +
		".page-break { page-break-before: always; }" +
		"</style></head><body>";

	html +=
		'<table class="footer-table"><tr>' +
		"<td><span>প্রস্তুতকারীর স্বাক্ষর</span></td>" +
		"<td><span>সংশ্লিষ্ট তত্বাবধায়কের স্বাক্ষর</span></td>" +
		"<td><span>হিসাব কর্মকর্তার স্বাক্ষর</span></td>" +
		"<td><span>অনুমোদনকারীর স্বাক্ষর</span></td>" +
		"</tr></table>";

	var MAX_PER_PAGE = 30;
	var numPages = Math.ceil(data.length / MAX_PER_PAGE) || 1;
	var rateNum = parseFloat(travelRate);

	for (var page = 0; page < numPages; page++) {
		if (page > 0) html += '<div class="page-break"></div>';

		html +=
			"<h1>ব্র্যাক - আড়ং</h1>" +
			"<h2>গড়পাড়া</h2>" +
			"<h3>পিছ রেট কর্মীদের মাসিক যাতায়াত বিল।</h3>" +
			'<div class="timestamp">Generated: ' +
			timestamp +
			"</div>" +
			'<table class="header-table"><tr>' +
			'<td style="width: 65%; border: none;">সেকশনঃ ' +
			sectionName +
			"</td>" +
			'<td style="width: 35%; border: none">মাসঃ ' +
			banglaMonthDisplay +
			"</td>" +
			"</tr></table>" +
			"<table><thead><tr>" +
			'<th style="width: 5%">ক্রঃ নং</th>' +
			'<th style="width: 23%">কর্মীর নাম</th>' +
			'<th style="width: 10%">কোড নং</th>' +
			'<th style="width: 12%">মোট উপস্থিত<br>কর্ম দিবস</th>' +
			'<th style="width: 15%">যাতায়াত ভাতার<br>হার</th>' +
			'<th style="width: 15%">মোট প্রাপ্ত<br>যাতায়াত ভাতা</th>' +
			'<th style="width: 20%">স্বাক্ষর</th>' +
			"</tr></thead><tbody>";

		var startIdx = page * MAX_PER_PAGE;
		var endIdx = Math.min(startIdx + MAX_PER_PAGE, data.length);
		var pageTotalAmount = 0;

		for (var i = startIdx; i < endIdx; i++) {
			var row = data[i];
			var days = parseInt(row.days) || 1;
			var totalPerArtisan = days * rateNum;
			pageTotalAmount += totalPerArtisan;

			html +=
				"<tr>" +
				"<td>" +
				toBanglaNum(i + 1) +
				"</td>" +
				'<td style="text-align: left;">' +
				row.name +
				"</td>" +
				"<td>" +
				toBanglaNum(row.code) +
				"</td>" +
				"<td>" +
				toBanglaNum(days) +
				"</td>" +
				"<td>" +
				banglaRate +
				"</td>" +
				"<td>" +
				toBanglaNum(totalPerArtisan) +
				"</td>" +
				"<td></td></tr>";
		}

		var banglaPageTotal = toBanglaNum(pageTotalAmount);
		var pageAmountInWords = numberToBanglaWords(pageTotalAmount);

		html +=
			'<tr style="font-weight:bold; background-color:#f8f8f8;">' +
			'<td colspan="5">মোট টাকাঃ</td>' +
			"<td>" +
			banglaPageTotal +
			"</td>" +
			"<td></td></tr>";

		html +=
			'<tr style="background-color:#f0f0f0;">' +
			'<td colspan="7" style="text-align:center; font-weight:bold; font-size:14px; padding:4px 8px; border-left:0.5px solid #8b8b8b; border-right:0.5px solid #8b8b8b; border-bottom:0.5px solid #8b8b8b;">' +
			"কথায়ঃ " +
			pageAmountInWords +
			" টাকা মাত্র।</td></tr>";

		html += "</tbody></table>";
	}

	html += "</body></html>";
	return html;
}
