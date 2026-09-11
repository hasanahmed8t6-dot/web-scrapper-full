var USPTO_API_KEY = "";

function doGet(e) {
  var serialNumber = (e && e.parameter && e.parameter.serial)
    ? String(e.parameter.serial).replace(/\D/g, "")
    : "";

  if (!serialNumber) {
    return jsonResponse({
      success: true,
      message: "Connected!"
    });
  }

  try {
    var result = scrapeTrademarkData(serialNumber);
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({
      success: false,
      serialNumber: serialNumber,
      error: err.toString()
    });
  }
}

function doPost(e) {
  return doGet(e);
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function scrapeTrademarkData(serialNumber) {
  var result = {
    success: true,
    serialNumber: serialNumber,

    // Existing fields — kept compatible with the current frontend.
    mark: "",
    correspondentName: "",
    correspondentAddress: "",
    phone: "",
    email: "",
    status: "",
    hasAttorney: false,
    attorneyName: "",

    // New fields.
    internationalClass: "",
    goodsServices: "",
    filingDate: "",
    statusDate: "",
    ownerName: "",
    ownerNames: [],
    ownerAddress: "",
    ownerAddresses: [],
    entityType: "",
    stateOrCountry: "",
    markType: "",
    register: "",
    basis: "",
    tm5Status: "",
    markDescription: "",
    disclaimer: "",

    confidence: "low",
    needsReview: false
  };

  try {
    // ------------------------------------------------------------
    // METHOD 1 — USPTO TSDR HTML
    // This is the primary method because it exposes the fields
    // visible on the actual TSDR status/document pages.
    // ------------------------------------------------------------
    var url = "https://tsdr.uspto.gov/statusview/sn" + serialNumber;
    var response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      deadline: 25,
      headers: {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.5",
        "Referer": "https://tsdr.uspto.gov/"
      }
    });

    if (response.getResponseCode() === 200) {
      var html = response.getContentText();

      function cleanHtml(value) {
        return String(value || "")
          .replace(/<br\s*\/?>/gi, "\n")
          .replace(/<script[\s\S]*?<\/script>/gi, "")
          .replace(/<style[\s\S]*?<\/style>/gi, "")
          .replace(/<[^>]+>/g, " ")
          .replace(/&amp;/g, "&")
          .replace(/&nbsp;/g, " ")
          .replace(/&#39;/g, "'")
          .replace(/&quot;/g, '"')
          .replace(/&#x27;/gi, "'")
          .replace(/&lt;/g, "<")
          .replace(/&gt;/g, ">")
          .replace(/\r/g, "")
          .replace(/[ \t]+/g, " ")
          .replace(/\n[ \t]+/g, "\n")
          .trim();
      }

      function extractField(key) {
        var escapedKey = key.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        var pattern = new RegExp(
          '<div[^>]*class="key"[^>]*>\\s*' + escapedKey +
          '\\s*<\\/div>\\s*' +
          '<div[^>]*class="value[^"]*"[^>]*>([\\s\\S]*?)<\\/div>',
          "i"
        );
        var match = html.match(pattern);
        return match && match[1] ? cleanHtml(match[1]).replace(/\s+/g, " ").trim() : "";
      }

      function extractAnyField(keys) {
        for (var i = 0; i < keys.length; i++) {
          var value = extractField(keys[i]);
          if (value) return value;
        }
        return "";
      }

      function firstMatch(patterns) {
        for (var i = 0; i < patterns.length; i++) {
          var m = html.match(patterns[i]);
          if (m && m[1]) {
            var value = cleanHtml(m[1]).trim();
            if (value) return value;
          }
        }
        return "";
      }

      // Basic existing fields.
      result.mark = extractAnyField(["Mark:"]) || result.mark;
      result.status = extractAnyField(["Status:"]) || result.status;
      result.phone = extractAnyField(["Phone:"]) || result.phone;

      // Additional mark/application information.
      result.internationalClass =
        extractAnyField([
          "International Class(es):",
          "International Class:",
          "International Class(es)"
        ]) || result.internationalClass;

      result.goodsServices =
        extractAnyField([
          "For:",
          "Goods and Services:",
          "Goods and Services"
        ]) || result.goodsServices;

      result.filingDate =
        extractAnyField([
          "Application Filing Date:",
          "Filing Date:",
          "Application Filing Date"
        ]) || result.filingDate;

      result.statusDate =
        extractAnyField(["Status Date:", "Status Date"]) || result.statusDate;

      result.markType =
        extractAnyField(["Mark Type:", "Mark Type"]) || result.markType;

      result.register =
        extractAnyField(["Register:", "Register"]) || result.register;

      result.tm5Status =
        extractAnyField([
          "TM5 Common Status Descriptor:",
          "TM5 Common Status Descriptor"
        ]) || result.tm5Status;

      result.markDescription =
        extractAnyField([
          "Description of Mark:",
          "Description of Mark"
        ]) || result.markDescription;

      result.disclaimer =
        extractAnyField(["Disclaimer:", "Disclaimer"]) || result.disclaimer;

      result.basis =
        extractAnyField([
          "Basis Information (Case Level):",
          "Basis Information (Case Level)"
        ]) || result.basis;

      // ------------------------------------------------------------
      // OWNER INFORMATION
      // TSDR can contain multiple Current Owner(s) blocks.
      // ------------------------------------------------------------
      var ownerBlockRegex =
        /Current\s+Owner(?:\(s\))?\s+Information[\s\S]*?(?=Attorney\/Correspondence\s+Information|Prosecution\s+History|TM\s+Staff|File\s+Location|$)/i;
      var ownerBlockMatch = html.match(ownerBlockRegex);

      if (ownerBlockMatch && ownerBlockMatch[0]) {
        var ownerBlock = ownerBlockMatch[0];

        var ownerMatches = [];
        var ownerRegex =
          /Owner\s+Name:\s*([\s\S]*?)(?=\s*Owner\s+Address:|\s*Legal\s+Entity\s+Type:|\s*State\s+or\s+Country\s+Where\s+Organized:|\s*$)/gi;
        var om;
        while ((om = ownerRegex.exec(ownerBlock)) !== null) {
          var owner = cleanHtml(om[1]).replace(/\s+/g, " ").trim();
          if (owner) ownerMatches.push(owner);
        }

        // Fallback for HTML where labels are div.key/value pairs.
        if (ownerMatches.length === 0) {
          var ownerKeyRegex =
            /<div[^>]*class="key"[^>]*>\s*Owner\s+Name:\s*<\/div>\s*<div[^>]*class="value[^"]*"[^>]*>([\s\S]*?)<\/div>/gi;
          while ((om = ownerKeyRegex.exec(ownerBlock)) !== null) {
            var owner2 = cleanHtml(om[1]).replace(/\s+/g, " ").trim();
            if (owner2) ownerMatches.push(owner2);
          }
        }

        result.ownerNames = ownerMatches;

        var ownerAddresses = [];
        var addressRegex =
          /Owner\s+Address:\s*([\s\S]*?)(?=\s*Legal\s+Entity\s+Type:|\s*State\s+or\s+Country\s+Where\s+Organized:|\s*Owner\s+Name:|\s*$)/gi;
        var adm;
        while ((adm = addressRegex.exec(ownerBlock)) !== null) {
          var addr = cleanHtml(adm[1]).replace(/\s+/g, " ").trim();
          if (addr) ownerAddresses.push(addr);
        }

        if (ownerAddresses.length === 0) {
          var ownerAddrKeyRegex =
            /<div[^>]*class="key"[^>]*>\s*Owner\s+Address:\s*<\/div>\s*<div[^>]*class="value[^"]*"[^>]*>([\s\S]*?)<\/div>/gi;
          while ((adm = ownerAddrKeyRegex.exec(ownerBlock)) !== null) {
            var addr2 = cleanHtml(adm[1]).replace(/\s+/g, " ").trim();
            if (addr2) ownerAddresses.push(addr2);
          }
        }

        result.ownerAddresses = ownerAddresses;

        var entityMatches = [];
        var entityRegex =
          /Legal\s+Entity\s+Type:\s*([\s\S]*?)(?=\s*State\s+or\s+Country\s+Where\s+Organized:|\s*Owner\s+Name:|\s*$)/gi;
        var em;
        while ((em = entityRegex.exec(ownerBlock)) !== null) {
          var entity = cleanHtml(em[1]).replace(/\s+/g, " ").trim();
          if (entity) entityMatches.push(entity);
        }

        if (entityMatches.length === 0) {
          var entityKeyRegex =
            /<div[^>]*class="key"[^>]*>\s*Legal\s+Entity\s+Type:\s*<\/div>\s*<div[^>]*class="value[^"]*"[^>]*>([\s\S]*?)<\/div>/gi;
          while ((em = entityKeyRegex.exec(ownerBlock)) !== null) {
            var entity2 = cleanHtml(em[1]).replace(/\s+/g, " ").trim();
            if (entity2) entityMatches.push(entity2);
          }
        }

        if (ownerMatches.length) result.ownerName = ownerMatches.join(" | ");
        if (ownerAddresses.length) result.ownerAddress = ownerAddresses.join(" | ");
        if (entityMatches.length) result.entityType = entityMatches.join(" | ");

        var stateMatches = [];
        var stateRegex =
          /State\s+or\s+Country\s+Where\s+Organized:\s*([\s\S]*?)(?=\s*Owner\s+Name:|\s*$)/gi;
        var sm;
        while ((sm = stateRegex.exec(ownerBlock)) !== null) {
          var state = cleanHtml(sm[1]).replace(/\s+/g, " ").trim();
          if (state) stateMatches.push(state);
        }
        if (stateMatches.length) result.stateOrCountry = stateMatches.join(" | ");
      }

      // If owner block parsing did not find anything, try global fields.
      if (!result.ownerName) {
        result.ownerName = firstMatch([
          /Owner\s+Name:\s*<\/?[^>]*>\s*([^<\n\r]+?)(?=\s*Owner\s+Address:)/i,
          /Owner\s+Name:\s*([^\n\r<]+)/i
        ]);
      }

      // ------------------------------------------------------------
      // ATTORNEY
      // ------------------------------------------------------------
      var attyOfRecord = extractAnyField(["Attorney Name:", "Attorney Name"]);
      if (html.indexOf("Attorney of Record - None") > -1) {
        result.hasAttorney = false;
        result.attorneyName = "";
        result.confidence = "high";
      } else if (attyOfRecord && attyOfRecord.length > 1) {
        result.hasAttorney = true;
        result.attorneyName = attyOfRecord;
        result.confidence = "high";
      } else {
        var attyFallback = firstMatch([
          /Attorney\s+of\s+Record\s*[-:]\s*([^<\n\r]{2,100})/i,
          /Attorney\s+Name:\s*([^<\n\r]{2,100})/i
        ]);
        if (attyFallback && attyFallback.toLowerCase() !== "none") {
          result.hasAttorney = true;
          result.attorneyName = attyFallback;
          result.confidence = "high";
        }
      }

      // ------------------------------------------------------------
      // EMAIL
      // ------------------------------------------------------------
      var emailPattern = /href=["']mailto:([^"']+)["']/i;
      var emailMatch = html.match(emailPattern);
      if (emailMatch && emailMatch[1]) {
        result.email = emailMatch[1].trim();
      } else {
        result.email = extractAnyField(["Correspondent e-mail:", "Correspondent e-mail"]);
      }

      // If multiple emails are printed together, normalize whitespace.
      result.email = result.email
        .replace(/\s+/g, " ")
        .trim();

      // ------------------------------------------------------------
      // CORRESPONDENT NAME + ADDRESS
      // ------------------------------------------------------------
      var corrPattern =
        /<div[^>]*class="key"[^>]*>\s*Correspondent\s+Name\/Address:\s*<\/div>\s*<div[^>]*class="value[^"]*"[^>]*>([\s\S]*?)<\/div>\s*(?=<div[^>]*class="(?:key|row))/i;
      var corrMatch = html.match(corrPattern);

      if (!corrMatch) {
        corrPattern =
          /<div[^>]*class="key"[^>]*>\s*Correspondent\s+Name\/Address:\s*<\/div>\s*<div[^>]*>([\s\S]*?)<\/div>\s*<div[^>]*class="key"/i;
        corrMatch = html.match(corrPattern);
      }

      if (!corrMatch) {
        var corrIndex = html.search(/Correspondent\s+Name\/Address:/i);
        if (corrIndex > -1) {
          var corrChunk = html.substring(corrIndex, corrIndex + 1200);
          var innerMatch = corrChunk.match(
            /<div[^>]*class="value[^"]*"[^>]*>([\s\S]*?)<\/div>/i
          );
          if (innerMatch) corrMatch = innerMatch;
        }
      }

      if (corrMatch && corrMatch[1]) {
        var corrRaw = corrMatch[1]
          .replace(/<br\s*\/?>/gi, "\n")
          .replace(/<[^>]+>/g, "")
          .replace(/&amp;/g, "&")
          .replace(/&nbsp;/g, " ")
          .replace(/&#39;/g, "'")
          .replace(/&quot;/g, '"')
          .trim();

        corrRaw = corrRaw.replace(/\n?Phone:[\s\S]*/i, "");
        corrRaw = corrRaw.replace(/\n?Correspondent\s+e-?mail:[\s\S]*/i, "");
        corrRaw = corrRaw.replace(/\n?E-?mail:[\s\S]*/i, "");
        corrRaw = corrRaw.trim();

        var corrLines = corrRaw
          .split("\n")
          .map(function(l) { return l.trim(); })
          .filter(function(l) { return l.length > 0; });

        if (corrLines.length > 0) result.correspondentName = corrLines[0];
        if (corrLines.length > 1) {
          result.correspondentAddress = corrLines.slice(1).join(", ");
        }
      }

      if (result.mark || result.status) result.confidence = "high";
      if (!result.mark && !result.status && !result.correspondentName) {
        result.needsReview = true;
        result.confidence = "low";
      }
    } else {
      result.needsReview = true;
      result.error = "USPTO TSDR returned HTTP " + response.getResponseCode();
    }

    // ------------------------------------------------------------
    // METHOD 2 — USPTO XML fallback
    // Only fills missing values.
    // ------------------------------------------------------------
    try {
      var xmlUrl = "https://tsdr.uspto.gov/documentparser/" + serialNumber + "/case.xml";
      var xmlResp = UrlFetchApp.fetch(xmlUrl, {
        muteHttpExceptions: true,
        followRedirects: true,
        deadline: 20
      });

      if (xmlResp.getResponseCode() === 200) {
        var x = xmlResp.getContentText();

        function xmlValue(patterns) {
          for (var i = 0; i < patterns.length; i++) {
            var m = x.match(patterns[i]);
            if (m && m[1]) return m[1].trim();
          }
          return "";
        }

        if (!result.mark) {
          result.mark = xmlValue([
            /<(?:[a-z0-9]+:)?MarkVerbalElementText>([^<]+)<\/(?:[a-z0-9]+:)?MarkVerbalElementText>/i,
            /<(?:[a-z0-9]+:)?markVerbalElementText>([^<]+)<\/(?:[a-z0-9]+:)?markVerbalElementText>/i
          ]);
        }

        if (!result.status) {
          result.status = xmlValue([
            /<(?:[a-z0-9]+:)?MarkCurrentStatusExternalDescriptionText>([^<]+)<\/(?:[a-z0-9]+:)?MarkCurrentStatusExternalDescriptionText>/i,
            /<(?:[a-z0-9]+:)?statusDescriptionText>([^<]+)<\/(?:[a-z0-9]+:)?statusDescriptionText>/i
          ]);
        }

        if (!result.correspondentName) {
          result.correspondentName = xmlValue([
            /<(?:[a-z0-9]+:)?corrAddr1>([^<]+)<\/(?:[a-z0-9]+:)?corrAddr1>/i
          ]);
        }

        if (!result.phone) {
          result.phone = xmlValue([
            /<(?:[a-z0-9]+:)?corrPhone>([^<]+)<\/(?:[a-z0-9]+:)?corrPhone>/i
          ]);
        }

        if (!result.email) {
          result.email = xmlValue([
            /<(?:[a-z0-9]+:)?corrEmail>([^<]+)<\/(?:[a-z0-9]+:)?corrEmail>/i,
            /<(?:[a-z0-9]+:)?attrneyPrimaryEmailAddr>([^<]+)<\/(?:[a-z0-9]+:)?attrneyPrimaryEmailAddr>/i
          ]);
        }

        if (!result.attorneyName) {
          var xmlAttorney = xmlValue([
            /<(?:[a-z0-9]+:)?attrneyNm>([^<]+)<\/(?:[a-z0-9]+:)?attrneyNm>/i,
            /<(?:[a-z0-9]+:)?AttorneyName>([^<]+)<\/(?:[a-z0-9]+:)?AttorneyName>/i,
            /<(?:[a-z0-9]+:)?attorneyName>([^<]+)<\/(?:[a-z0-9]+:)?attorneyName>/i
          ]);
          if (xmlAttorney) {
            result.hasAttorney = true;
            result.attorneyName = xmlAttorney;
            result.confidence = "high";
          }
        }

        if (!result.internationalClass) {
          result.internationalClass = xmlValue([
            /<(?:[a-z0-9]+:)?internationalClass[^>]*>([^<]+)<\/(?:[a-z0-9]+:)?internationalClass>/i,
            /<(?:[a-z0-9]+:)?InternationalClass[^>]*>([^<]+)<\/(?:[a-z0-9]+:)?InternationalClass>/i
          ]);
        }
      }
    } catch (xmlErr) {
      Logger.log("XML fallback failed: " + xmlErr.toString());
    }

    // ------------------------------------------------------------
    // METHOD 3 — Status JSON fallback
    // ------------------------------------------------------------
    try {
      var jsonUrl = "https://tsdrapi.uspto.gov/ts/cd/casestatus/" + serialNumber + "/info.json";
      var jsonResp = UrlFetchApp.fetch(jsonUrl, {
        muteHttpExceptions: true,
        deadline: 15,
        headers: { "Accept": "application/json" }
      });

      if (jsonResp.getResponseCode() === 200) {
        var jt = jsonResp.getContentText();

        function jsonValue(patterns) {
          for (var i = 0; i < patterns.length; i++) {
            var m = jt.match(patterns[i]);
            if (m && m[1] && m[1].toLowerCase() !== "null") return m[1].trim();
          }
          return "";
        }

        if (!result.mark) {
          result.mark = jsonValue([
            /"markVerbalElementText"\s*:\s*"([^"]+)"/i,
            /"MarkVerbalElementText"\s*:\s*"([^"]+)"/i,
            /"wordMark"\s*:\s*"([^"]+)"/i
          ]);
        }

        if (!result.status) {
          result.status = jsonValue([
            /"markCurrentStatusExternalDescriptionText"\s*:\s*"([^"]+)"/i,
            /"statusDescriptionText"\s*:\s*"([^"]+)"/i
          ]);
        }

        if (!result.correspondentName) {
          result.correspondentName = jsonValue([
            /"corrAddr1"\s*:\s*"([^"]+)"/i,
            /"correspondentName"\s*:\s*"([^"]+)"/i
          ]);
        }

        if (!result.phone) {
          result.phone = jsonValue([
            /"corrPhone"\s*:\s*"([^"]+)"/i,
            /"phone"\s*:\s*"([^"]+)"/i
          ]);
        }

        if (!result.email) {
          result.email = jsonValue([
            /"corrEmail"\s*:\s*"([^"]+)"/i,
            /"attrneyPrimaryEmailAddr"\s*:\s*"([^"]+)"/i,
            /"primaryEmailAddr"\s*:\s*"([^"]+)"/i
          ]);
        }

        if (!result.attorneyName) {
          var jsonAttorney = jsonValue([
            /"attrneyNm"\s*:\s*"([^"]+)"/i,
            /"attorneyName"\s*:\s*"([^"]+)"/i,
            /"AttorneyName"\s*:\s*"([^"]+)"/i,
            /"attrnyNm"\s*:\s*"([^"]+)"/i
          ]);
          if (jsonAttorney) {
            result.hasAttorney = true;
            result.attorneyName = jsonAttorney;
            result.confidence = "high";
          }
        }
      }
    } catch (jsonErr) {
      Logger.log("JSON fallback failed: " + jsonErr.toString());
    }

    // Final state.
    if (result.ownerNames.length && !result.ownerName) {
      result.ownerName = result.ownerNames.join(" | ");
    }

    if (result.ownerAddresses.length && !result.ownerAddress) {
      result.ownerAddress = result.ownerAddresses.join(" | ");
    }

    if (result.mark || result.status || result.ownerName) {
      if (result.confidence === "low") result.confidence = "medium";
    }

    if (!result.mark && !result.status && !result.correspondentName && !result.ownerName) {
      result.needsReview = true;
      result.confidence = "low";
    }

  } catch (err) {
    result.needsReview = true;
    result.success = false;
    result.error = err.toString();
  }

  return result;
}

function testSerial() {
  var r = scrapeTrademarkData("50038347");
  Logger.log(JSON.stringify(r, null, 2));
}
