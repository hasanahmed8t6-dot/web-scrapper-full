import { useEffect, useRef, useState } from "react";

interface TrademarkResult {
  serialNumber: string;
  mark: string;
  status: string;
  correspondentName: string;
  correspondentAddress: string;
  phone: string;
  email: string;
  hasAttorney: boolean;
  attorneyName: string;
  confidence: string;
  needsReview: boolean;
  internationalClass: string;
  goodsServices: string;
  filingDate: string;
  statusDate: string;
  ownerName: string;
  ownerAddress: string;
  entityType: string;
  stateOrCountry: string;
  markType: string;
  register: string;
  basis: string;
  tm5Status: string;
  markDescription: string;
  disclaimer: string;
  error?: string;
  pending?: boolean;
}

type TabType = "all" | "attorney" | "noAttorney" | "errors";

const APPS_SCRIPT_CODE = `var USPTO_API_KEY = "";

function doGet(e) {
  var serialNumber = (e && e.parameter && e.parameter.serial)
    ? String(e.parameter.serial).replace(/\\D/g, "")
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
          .replace(/<br\\s*\\/?>/gi, "\\n")
          .replace(/<script[\\s\\S]*?<\\/script>/gi, "")
          .replace(/<style[\\s\\S]*?<\\/style>/gi, "")
          .replace(/<[^>]+>/g, " ")
          .replace(/&amp;/g, "&")
          .replace(/&nbsp;/g, " ")
          .replace(/&#39;/g, "'")
          .replace(/&quot;/g, '"')
          .replace(/&#x27;/gi, "'")
          .replace(/&lt;/g, "<")
          .replace(/&gt;/g, ">")
          .replace(/\\r/g, "")
          .replace(/[ \\t]+/g, " ")
          .replace(/\\n[ \\t]+/g, "\\n")
          .trim();
      }

      function extractField(key) {
        var escapedKey = key.replace(/[.*+?^\${}()|[\\]\\\\]/g, "\\\\$&");
        var pattern = new RegExp(
          '<div[^>]*class="key"[^>]*>\\\\s*' + escapedKey +
          '\\\\s*<\\\\/div>\\\\s*' +
          '<div[^>]*class="value[^"]*"[^>]*>([\\\\s\\\\S]*?)<\\\\/div>',
          "i"
        );
        var match = html.match(pattern);
        return match && match[1] ? cleanHtml(match[1]).replace(/\\s+/g, " ").trim() : "";
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
        /Current\\s+Owner(?:\\(s\\))?\\s+Information[\\s\\S]*?(?=Attorney\\/Correspondence\\s+Information|Prosecution\\s+History|TM\\s+Staff|File\\s+Location|$)/i;
      var ownerBlockMatch = html.match(ownerBlockRegex);

      if (ownerBlockMatch && ownerBlockMatch[0]) {
        var ownerBlock = ownerBlockMatch[0];

        var ownerMatches = [];
        var ownerRegex =
          /Owner\\s+Name:\\s*([\\s\\S]*?)(?=\\s*Owner\\s+Address:|\\s*Legal\\s+Entity\\s+Type:|\\s*State\\s+or\\s+Country\\s+Where\\s+Organized:|\\s*$)/gi;
        var om;
        while ((om = ownerRegex.exec(ownerBlock)) !== null) {
          var owner = cleanHtml(om[1]).replace(/\\s+/g, " ").trim();
          if (owner) ownerMatches.push(owner);
        }

        // Fallback for HTML where labels are div.key/value pairs.
        if (ownerMatches.length === 0) {
          var ownerKeyRegex =
            /<div[^>]*class="key"[^>]*>\\s*Owner\\s+Name:\\s*<\\/div>\\s*<div[^>]*class="value[^"]*"[^>]*>([\\s\\S]*?)<\\/div>/gi;
          while ((om = ownerKeyRegex.exec(ownerBlock)) !== null) {
            var owner2 = cleanHtml(om[1]).replace(/\\s+/g, " ").trim();
            if (owner2) ownerMatches.push(owner2);
          }
        }

        result.ownerNames = ownerMatches;

        var ownerAddresses = [];
        var addressRegex =
          /Owner\\s+Address:\\s*([\\s\\S]*?)(?=\\s*Legal\\s+Entity\\s+Type:|\\s*State\\s+or\\s+Country\\s+Where\\s+Organized:|\\s*Owner\\s+Name:|\\s*$)/gi;
        var adm;
        while ((adm = addressRegex.exec(ownerBlock)) !== null) {
          var addr = cleanHtml(adm[1]).replace(/\\s+/g, " ").trim();
          if (addr) ownerAddresses.push(addr);
        }

        if (ownerAddresses.length === 0) {
          var ownerAddrKeyRegex =
            /<div[^>]*class="key"[^>]*>\\s*Owner\\s+Address:\\s*<\\/div>\\s*<div[^>]*class="value[^"]*"[^>]*>([\\s\\S]*?)<\\/div>/gi;
          while ((adm = ownerAddrKeyRegex.exec(ownerBlock)) !== null) {
            var addr2 = cleanHtml(adm[1]).replace(/\\s+/g, " ").trim();
            if (addr2) ownerAddresses.push(addr2);
          }
        }

        result.ownerAddresses = ownerAddresses;

        var entityMatches = [];
        var entityRegex =
          /Legal\\s+Entity\\s+Type:\\s*([\\s\\S]*?)(?=\\s*State\\s+or\\s+Country\\s+Where\\s+Organized:|\\s*Owner\\s+Name:|\\s*$)/gi;
        var em;
        while ((em = entityRegex.exec(ownerBlock)) !== null) {
          var entity = cleanHtml(em[1]).replace(/\\s+/g, " ").trim();
          if (entity) entityMatches.push(entity);
        }

        if (entityMatches.length === 0) {
          var entityKeyRegex =
            /<div[^>]*class="key"[^>]*>\\s*Legal\\s+Entity\\s+Type:\\s*<\\/div>\\s*<div[^>]*class="value[^"]*"[^>]*>([\\s\\S]*?)<\\/div>/gi;
          while ((em = entityKeyRegex.exec(ownerBlock)) !== null) {
            var entity2 = cleanHtml(em[1]).replace(/\\s+/g, " ").trim();
            if (entity2) entityMatches.push(entity2);
          }
        }

        if (ownerMatches.length) result.ownerName = ownerMatches.join(" | ");
        if (ownerAddresses.length) result.ownerAddress = ownerAddresses.join(" | ");
        if (entityMatches.length) result.entityType = entityMatches.join(" | ");

        var stateMatches = [];
        var stateRegex =
          /State\\s+or\\s+Country\\s+Where\\s+Organized:\\s*([\\s\\S]*?)(?=\\s*Owner\\s+Name:|\\s*$)/gi;
        var sm;
        while ((sm = stateRegex.exec(ownerBlock)) !== null) {
          var state = cleanHtml(sm[1]).replace(/\\s+/g, " ").trim();
          if (state) stateMatches.push(state);
        }
        if (stateMatches.length) result.stateOrCountry = stateMatches.join(" | ");
      }

      // If owner block parsing did not find anything, try global fields.
      if (!result.ownerName) {
        result.ownerName = firstMatch([
          /Owner\\s+Name:\\s*<\\/?[^>]*>\\s*([^<\\n\\r]+?)(?=\\s*Owner\\s+Address:)/i,
          /Owner\\s+Name:\\s*([^\\n\\r<]+)/i
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
          /Attorney\\s+of\\s+Record\\s*[-:]\\s*([^<\\n\\r]{2,100})/i,
          /Attorney\\s+Name:\\s*([^<\\n\\r]{2,100})/i
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
        .replace(/\\s+/g, " ")
        .trim();

      // ------------------------------------------------------------
      // CORRESPONDENT NAME + ADDRESS
      // ------------------------------------------------------------
      var corrPattern =
        /<div[^>]*class="key"[^>]*>\\s*Correspondent\\s+Name\\/Address:\\s*<\\/div>\\s*<div[^>]*class="value[^"]*"[^>]*>([\\s\\S]*?)<\\/div>\\s*(?=<div[^>]*class="(?:key|row))/i;
      var corrMatch = html.match(corrPattern);

      if (!corrMatch) {
        corrPattern =
          /<div[^>]*class="key"[^>]*>\\s*Correspondent\\s+Name\\/Address:\\s*<\\/div>\\s*<div[^>]*>([\\s\\S]*?)<\\/div>\\s*<div[^>]*class="key"/i;
        corrMatch = html.match(corrPattern);
      }

      if (!corrMatch) {
        var corrIndex = html.search(/Correspondent\\s+Name\\/Address:/i);
        if (corrIndex > -1) {
          var corrChunk = html.substring(corrIndex, corrIndex + 1200);
          var innerMatch = corrChunk.match(
            /<div[^>]*class="value[^"]*"[^>]*>([\\s\\S]*?)<\\/div>/i
          );
          if (innerMatch) corrMatch = innerMatch;
        }
      }

      if (corrMatch && corrMatch[1]) {
        var corrRaw = corrMatch[1]
          .replace(/<br\\s*\\/?>/gi, "\\n")
          .replace(/<[^>]+>/g, "")
          .replace(/&amp;/g, "&")
          .replace(/&nbsp;/g, " ")
          .replace(/&#39;/g, "'")
          .replace(/&quot;/g, '"')
          .trim();

        corrRaw = corrRaw.replace(/\\n?Phone:[\\s\\S]*/i, "");
        corrRaw = corrRaw.replace(/\\n?Correspondent\\s+e-?mail:[\\s\\S]*/i, "");
        corrRaw = corrRaw.replace(/\\n?E-?mail:[\\s\\S]*/i, "");
        corrRaw = corrRaw.trim();

        var corrLines = corrRaw
          .split("\\n")
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
            /<(?:[a-z0-9]+:)?MarkVerbalElementText>([^<]+)<\\/(?:[a-z0-9]+:)?MarkVerbalElementText>/i,
            /<(?:[a-z0-9]+:)?markVerbalElementText>([^<]+)<\\/(?:[a-z0-9]+:)?markVerbalElementText>/i
          ]);
        }

        if (!result.status) {
          result.status = xmlValue([
            /<(?:[a-z0-9]+:)?MarkCurrentStatusExternalDescriptionText>([^<]+)<\\/(?:[a-z0-9]+:)?MarkCurrentStatusExternalDescriptionText>/i,
            /<(?:[a-z0-9]+:)?statusDescriptionText>([^<]+)<\\/(?:[a-z0-9]+:)?statusDescriptionText>/i
          ]);
        }

        if (!result.correspondentName) {
          result.correspondentName = xmlValue([
            /<(?:[a-z0-9]+:)?corrAddr1>([^<]+)<\\/(?:[a-z0-9]+:)?corrAddr1>/i
          ]);
        }

        if (!result.phone) {
          result.phone = xmlValue([
            /<(?:[a-z0-9]+:)?corrPhone>([^<]+)<\\/(?:[a-z0-9]+:)?corrPhone>/i
          ]);
        }

        if (!result.email) {
          result.email = xmlValue([
            /<(?:[a-z0-9]+:)?corrEmail>([^<]+)<\\/(?:[a-z0-9]+:)?corrEmail>/i,
            /<(?:[a-z0-9]+:)?attrneyPrimaryEmailAddr>([^<]+)<\\/(?:[a-z0-9]+:)?attrneyPrimaryEmailAddr>/i
          ]);
        }

        if (!result.attorneyName) {
          var xmlAttorney = xmlValue([
            /<(?:[a-z0-9]+:)?attrneyNm>([^<]+)<\\/(?:[a-z0-9]+:)?attrneyNm>/i,
            /<(?:[a-z0-9]+:)?AttorneyName>([^<]+)<\\/(?:[a-z0-9]+:)?AttorneyName>/i,
            /<(?:[a-z0-9]+:)?attorneyName>([^<]+)<\\/(?:[a-z0-9]+:)?attorneyName>/i
          ]);
          if (xmlAttorney) {
            result.hasAttorney = true;
            result.attorneyName = xmlAttorney;
            result.confidence = "high";
          }
        }

        if (!result.internationalClass) {
          result.internationalClass = xmlValue([
            /<(?:[a-z0-9]+:)?internationalClass[^>]*>([^<]+)<\\/(?:[a-z0-9]+:)?internationalClass>/i,
            /<(?:[a-z0-9]+:)?InternationalClass[^>]*>([^<]+)<\\/(?:[a-z0-9]+:)?InternationalClass>/i
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
            /"markVerbalElementText"\\s*:\\s*"([^"]+)"/i,
            /"MarkVerbalElementText"\\s*:\\s*"([^"]+)"/i,
            /"wordMark"\\s*:\\s*"([^"]+)"/i
          ]);
        }

        if (!result.status) {
          result.status = jsonValue([
            /"markCurrentStatusExternalDescriptionText"\\s*:\\s*"([^"]+)"/i,
            /"statusDescriptionText"\\s*:\\s*"([^"]+)"/i
          ]);
        }

        if (!result.correspondentName) {
          result.correspondentName = jsonValue([
            /"corrAddr1"\\s*:\\s*"([^"]+)"/i,
            /"correspondentName"\\s*:\\s*"([^"]+)"/i
          ]);
        }

        if (!result.phone) {
          result.phone = jsonValue([
            /"corrPhone"\\s*:\\s*"([^"]+)"/i,
            /"phone"\\s*:\\s*"([^"]+)"/i
          ]);
        }

        if (!result.email) {
          result.email = jsonValue([
            /"corrEmail"\\s*:\\s*"([^"]+)"/i,
            /"attrneyPrimaryEmailAddr"\\s*:\\s*"([^"]+)"/i,
            /"primaryEmailAddr"\\s*:\\s*"([^"]+)"/i
          ]);
        }

        if (!result.attorneyName) {
          var jsonAttorney = jsonValue([
            /"attrneyNm"\\s*:\\s*"([^"]+)"/i,
            /"attorneyName"\\s*:\\s*"([^"]+)"/i,
            /"AttorneyName"\\s*:\\s*"([^"]+)"/i,
            /"attrnyNm"\\s*:\\s*"([^"]+)"/i
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
`;

const getTsdrLink = (serial: string) =>
  `https://tsdr.uspto.gov/#caseNumber=${serial}&caseSearchType=US_APPLICATION&caseType=DEFAULT&searchType=statusSearch`;

const parseSerials = (input: string): string[] =>
  Array.from(
    new Set(
      input
        .split(/[\\s,;\\n]+/)
        .map((s) => s.trim().replace(/\\D/g, ""))
        .filter((s) => s.length >= 7 && s.length <= 9)
    )
  );

function Badge({ value }: { value: string | boolean | null }) {
  if (value === true || value === "true")
    return <span className="inline-flex rounded-full bg-emerald-500/20 px-2 py-0.5 text-xs font-semibold text-emerald-400 ring-1 ring-emerald-500/40">✓ Yes</span>;
  if (value === false || value === "false")
    return <span className="inline-flex rounded-full bg-amber-500/20 px-2 py-0.5 text-xs font-semibold text-amber-400 ring-1 ring-amber-500/40">✗ No</span>;
  return <span className="inline-flex rounded-full bg-slate-500/20 px-2 py-0.5 text-xs font-semibold text-slate-400 ring-1 ring-slate-500/40">? Unknown</span>;
}

function ConfidencePill({ level }: { level: string }) {
  const map: Record<string, string> = {
    high: "bg-emerald-500/20 text-emerald-400 ring-emerald-500/40",
    medium: "bg-yellow-500/20 text-yellow-400 ring-yellow-500/40",
    low: "bg-red-500/20 text-red-400 ring-red-500/40",
  };
  return <span className={`inline-flex rounded-full px-2 py-0.5 text-xs font-semibold ring-1 ${map[level] ?? map.low}`}>{level}</span>;
}

function CopyButton({ text, label }: { text: string; label: string }) {
  const [copied, setCopied] = useState(false);
  const copy = async () => {
    await navigator.clipboard.writeText(text);
    setCopied(true);
    setTimeout(() => setCopied(false), 1800);
  };
  return (
    <button onClick={copy} className="rounded-lg bg-blue-600 px-4 py-2 text-sm font-medium text-white hover:bg-blue-500">
      {copied ? "✅ Copied!" : `📋 ${label}`}
    </button>
  );
}

function cleanCell(value: unknown): string {
  return String(value ?? "").replace(/[\r\n]+/g, " ").trim();
}

export default function App() {
  const [step, setStep] = useState<1 | 2>(1);
  const [scriptUrl, setScriptUrl] = useState("");
  const [isTesting, setIsTesting] = useState(false);
  const [connectionError, setConnectionError] = useState("");
  const [serialInput, setSerialInput] = useState("");
  const [results, setResults] = useState<TrademarkResult[]>([]);
  const [isChecking, setIsChecking] = useState(false);
  const [progress, setProgress] = useState(0);
  const [currentSerial, setCurrentSerial] = useState("");
  const [activeTab, setActiveTab] = useState<TabType>("all");
  const [searchFilter, setSearchFilter] = useState("");
  const abortRef = useRef(false);

  useEffect(() => {
    const saved = localStorage.getItem("ladley_script_url");
    if (saved) {
      setScriptUrl(saved);
      setStep(2);
    }
  }, []);

  const testConnection = async () => {
    if (!scriptUrl.trim()) {
      setConnectionError("Please enter your Google Apps Script Web App URL.");
      return;
    }
    setIsTesting(true);
    setConnectionError("");
    try {
      const res = await fetch(`${scriptUrl.trim()}?serial=50038347`);
      const data = await res.json();
      if (data.success !== undefined || data.serialNumber !== undefined) {
        localStorage.setItem("ladley_script_url", scriptUrl.trim());
        setStep(2);
      } else {
        setConnectionError("Script responded but returned unexpected data.");
      }
    } catch (err: unknown) {
      setConnectionError("Connection failed: " + (err instanceof Error ? err.message : String(err)));
    } finally {
      setIsTesting(false);
    }
  };

  const runLookup = async () => {
    const serials = parseSerials(serialInput);
    if (!serials.length) {
      alert("Please enter at least one valid serial number (7–9 digits).");
      return;
    }

    abortRef.current = false;
    setIsChecking(true);
    setProgress(0);
    setResults([]);
    setActiveTab("all");

    const blank = (serial: string): TrademarkResult => ({
      serialNumber: serial,
      mark: "",
      status: "",
      correspondentName: "",
      correspondentAddress: "",
      phone: "",
      email: "",
      hasAttorney: false,
      attorneyName: "",
      confidence: "low",
      needsReview: false,
      internationalClass: "",
      goodsServices: "",
      filingDate: "",
      statusDate: "",
      ownerName: "",
      ownerAddress: "",
      entityType: "",
      stateOrCountry: "",
      markType: "",
      register: "",
      basis: "",
      tm5Status: "",
      markDescription: "",
      disclaimer: "",
      pending: true,
    });

    const updated = serials.map(blank);
    setResults([...updated]);

    for (let i = 0; i < serials.length; i++) {
      if (abortRef.current) break;

      const serial = serials[i];
      setCurrentSerial(serial);
      setProgress(Math.round(((i + 1) / serials.length) * 100));

      try {
        const res = await fetch(`${scriptUrl.trim()}?serial=${serial}`);
        const data = await res.json();

        updated[i] = {
          ...updated[i],
          serialNumber: serial,
          mark: data.mark || "",
          status: data.status || "",
          correspondentName: data.correspondentName || "",
          correspondentAddress: data.correspondentAddress || "",
          phone: data.phone || "",
          email: data.email || "",
          hasAttorney: !!data.hasAttorney,
          attorneyName: data.attorneyName || "",
          confidence: data.confidence || "low",
          needsReview: !!data.needsReview,
          internationalClass: data.internationalClass || "",
          goodsServices: data.goodsServices || "",
          filingDate: data.filingDate || "",
          statusDate: data.statusDate || "",
          ownerName: data.ownerName || (Array.isArray(data.ownerNames) ? data.ownerNames.join(" | ") : ""),
          ownerAddress: data.ownerAddress || (Array.isArray(data.ownerAddresses) ? data.ownerAddresses.join(" | ") : ""),
          entityType: data.entityType || "",
          stateOrCountry: data.stateOrCountry || "",
          markType: data.markType || "",
          register: data.register || "",
          basis: data.basis || "",
          tm5Status: data.tm5Status || "",
          markDescription: data.markDescription || "",
          disclaimer: data.disclaimer || "",
          error: data.error || undefined,
          pending: false,
        };
      } catch (err: unknown) {
        updated[i] = {
          ...updated[i],
          error: err instanceof Error ? err.message : String(err),
          pending: false,
        };
      }

      setResults([...updated]);

      if (i < serials.length - 1 && !abortRef.current) {
        await new Promise((resolve) => setTimeout(resolve, 600));
      }
    }

    setIsChecking(false);
    setCurrentSerial("");
  };

  const stopLookup = () => {
    abortRef.current = true;
    setIsChecking(false);
    setCurrentSerial("");
  };

  const exportCSV = (rows: TrademarkResult[], filename: string) => {
    if (!rows.length) {
      alert("No data to export.");
      return;
    }

    const headers = [
      "Serial Number", "Mark", "Status", "International Class", "Goods/Services",
      "Filing Date", "Status Date", "Owner Name", "Owner Address", "Entity Type",
      "State/Country Organized", "Mark Type", "Register", "Basis", "TM5 Status",
      "Mark Description", "Disclaimer", "Has Attorney", "Attorney Name",
      "Correspondent Name", "Correspondent Address", "Phone", "Correspondent Email",
      "Confidence", "Needs Review", "TSDR Link"
    ];

    const escape = (v: unknown) => `"${cleanCell(v).replace(/"/g, '""')}"`;

    const lines = rows.map((r) =>
      [
        escape(r.serialNumber), escape(r.mark), escape(r.status), escape(r.internationalClass),
        escape(r.goodsServices), escape(r.filingDate), escape(r.statusDate), escape(r.ownerName),
        escape(r.ownerAddress), escape(r.entityType), escape(r.stateOrCountry), escape(r.markType),
        escape(r.register), escape(r.basis), escape(r.tm5Status), escape(r.markDescription),
        escape(r.disclaimer), r.hasAttorney ? "Yes" : "No", escape(r.attorneyName),
        escape(r.correspondentName), escape(r.correspondentAddress), escape(r.phone), escape(r.email),
        escape(r.confidence), r.needsReview ? "Yes" : "No", escape(getTsdrLink(r.serialNumber))
      ].join(",")
    );

    const csv = "\\ufeff" + [headers.join(","), ...lines].join("\\r\\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 100);
  };

  const withAttorney = results.filter((r) => !r.pending && !r.error && r.hasAttorney);
  const noAttorney = results.filter((r) => !r.pending && !r.error && !r.hasAttorney);
  const errors = results.filter((r) => !r.pending && !!r.error);
  const filterText = searchFilter.toLowerCase();

  const baseRows =
    activeTab === "attorney" ? withAttorney :
    activeTab === "noAttorney" ? noAttorney :
    activeTab === "errors" ? errors :
    results.filter((r) => !r.pending && !r.error);

  const filteredRows = baseRows.filter((r) =>
    !filterText ||
    r.serialNumber.includes(filterText) ||
    r.mark.toLowerCase().includes(filterText) ||
    r.ownerName.toLowerCase().includes(filterText) ||
    r.correspondentName.toLowerCase().includes(filterText) ||
    r.email.toLowerCase().includes(filterText) ||
    r.status.toLowerCase().includes(filterText) ||
    r.goodsServices.toLowerCase().includes(filterText) ||
    r.internationalClass.toLowerCase().includes(filterText)
  );

  const tabClass = (t: TabType) =>
    `rounded-t-lg px-4 py-2 text-sm font-medium transition-colors ${
      activeTab === t ? "border-b-2 border-blue-500 bg-slate-800 text-white" : "bg-slate-900 text-slate-400 hover:text-white"
    }`;

  if (step === 1) {
    return (
      <div className="min-h-screen bg-gradient-to-br from-slate-950 via-slate-900 to-slate-950 text-white">
        <div className="mx-auto max-w-3xl px-4 py-10">
          <div className="mb-10 text-center">
            <div className="mb-3 inline-flex rounded-full border border-blue-500/30 bg-blue-500/10 px-4 py-1.5 text-sm text-blue-400">⚖️ USPTO Trademark Intelligence</div>
            <h1 className="text-4xl font-extrabold">LADLEY Trademark Lookup</h1>
            <p className="mt-2 text-slate-400">Bulk-check USPTO serial numbers for mark, owner, goods/services, contact information and attorney status.</p>
          </div>

          <div className="rounded-2xl border border-slate-700 bg-slate-800/60 p-8 shadow-2xl">
            <h2 className="mb-6 text-xl font-bold text-yellow-400">⚙️ One-Time Setup</h2>

            <div className="space-y-6">
              <div className="rounded-xl border border-slate-600 bg-slate-900/60 p-5">
                <h3 className="mb-3 font-bold text-blue-300">1. Google Apps Script</h3>
                <p className="mb-4 text-sm text-slate-300">
                  Use the updated <strong>google-apps-script/Code.gs</strong> included in this repository. It preserves your existing fields and adds owner, class, goods/services, filing date and other TSDR information.
                </p>
                <CopyButton text={APPS_SCRIPT_CODE} label="Copy Apps Script Code" />
              </div>

              <div className="rounded-xl border border-slate-600 bg-slate-900/60 p-5">
                <h3 className="mb-3 font-bold text-blue-300">2. Deploy as Web App</h3>
                <ol className="list-decimal space-y-1.5 pl-5 text-sm text-slate-300">
                  <li>Open script.google.com and open your existing scraper project.</li>
                  <li>Replace its code with <strong>Code.gs</strong>.</li>
                  <li>Save the project.</li>
                  <li>Deploy → Manage deployments → Edit the existing Web App.</li>
                  <li>Set access to <strong className="text-yellow-400">Anyone</strong>.</li>
                  <li>Deploy the new version.</li>
                </ol>
              </div>

              <div className="rounded-xl border border-slate-600 bg-slate-900/60 p-5">
                <h3 className="mb-3 font-bold text-blue-300">3. Connect</h3>
                <input
                  value={scriptUrl}
                  onChange={(e) => setScriptUrl(e.target.value)}
                  placeholder="https://script.google.com/macros/s/.../exec"
                  className="mb-3 w-full rounded-lg border border-slate-600 bg-slate-800 px-4 py-3 text-sm text-white placeholder-slate-500 focus:border-blue-500 focus:outline-none"
                />
                {connectionError && <div className="mb-3 rounded-lg border border-red-500/40 bg-red-500/10 px-4 py-2 text-sm text-red-400">❌ {connectionError}</div>}
                <button
                  onClick={testConnection}
                  disabled={isTesting || !scriptUrl.trim()}
                  className="rounded-lg bg-emerald-600 px-6 py-2.5 text-sm font-bold text-white hover:bg-emerald-500 disabled:opacity-50"
                >
                  {isTesting ? "Testing…" : "🔗 Test Connection & Continue"}
                </button>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-950 via-slate-900 to-slate-950 text-white">
      <div className="mx-auto max-w-[1800px] px-4 py-8">
        <div className="mb-8 flex flex-wrap items-center justify-between gap-4">
          <div>
            <h1 className="text-3xl font-extrabold">⚖️ LADLEY Trademark Lookup</h1>
            <p className="text-sm text-slate-400">USPTO TSDR Bulk Intelligence — Mark · Owner · Goods/Services · Class · Contact · Attorney</p>
          </div>
          <div className="flex items-center gap-3">
            <span className="flex items-center gap-1.5 rounded-full border border-emerald-500/30 bg-emerald-500/10 px-3 py-1 text-xs text-emerald-400"><span className="h-2 w-2 animate-pulse rounded-full bg-emerald-400" /> Connected</span>
            <button onClick={() => { localStorage.removeItem("ladley_script_url"); setStep(1); setResults([]); }} className="rounded-lg border border-slate-600 px-3 py-1.5 text-xs text-slate-400 hover:text-white">⚙️ Change Script</button>
          </div>
        </div>

        <div className="mb-6 rounded-2xl border border-slate-700 bg-slate-800/60 p-6 shadow-xl">
          <label className="mb-2 block text-sm font-semibold text-slate-300">USPTO Serial Numbers</label>
          <textarea
            value={serialInput}
            onChange={(e) => setSerialInput(e.target.value)}
            placeholder={"Paste serial numbers separated by commas, spaces, or new lines.\\nExample:\\n50038347\\n99857081"}
            rows={5}
            className="w-full resize-none rounded-xl border border-slate-600 bg-slate-900 px-4 py-3 font-mono text-sm text-white placeholder-slate-500 focus:border-blue-500 focus:outline-none"
          />
          <div className="mt-4 flex flex-wrap items-center gap-3">
            {!isChecking ? (
              <button onClick={runLookup} disabled={!serialInput.trim()} className="rounded-xl bg-blue-600 px-6 py-3 text-sm font-bold hover:bg-blue-500 disabled:opacity-50">🔍 Lookup Trademarks</button>
            ) : (
              <button onClick={stopLookup} className="rounded-xl bg-red-600 px-6 py-3 text-sm font-bold hover:bg-red-500">⏹ Stop</button>
            )}
            {results.length > 0 && !isChecking && <button onClick={() => { setResults([]); setSerialInput(""); }} className="rounded-xl border border-slate-600 px-4 py-3 text-sm text-slate-400 hover:text-white">Clear Results</button>}
            {serialInput.trim() && <span className="text-sm text-slate-500">{parseSerials(serialInput).length} serial number(s) detected</span>}
          </div>
        </div>

        {isChecking && (
          <div className="mb-6 rounded-2xl border border-slate-700 bg-slate-800/60 p-5">
            <div className="mb-2 flex justify-between text-sm"><span className="text-slate-300">Checking <span className="font-mono text-blue-400">{currentSerial}</span>…</span><span className="text-slate-400">{progress}%</span></div>
            <div className="h-2.5 w-full overflow-hidden rounded-full bg-slate-700"><div className="h-full bg-gradient-to-r from-blue-500 to-violet-500 transition-all" style={{ width: `${progress}%` }} /></div>
          </div>
        )}

        {results.some((r) => !r.pending) && (
          <div className="mb-6 grid grid-cols-2 gap-4 sm:grid-cols-4">
            <div className="rounded-2xl border border-slate-700 bg-slate-800/60 p-4 text-center"><div className="text-3xl font-black">{results.filter((r) => !r.pending && !r.error).length}</div><div className="text-xs text-slate-400">Total Checked</div></div>
            <div className="rounded-2xl border border-emerald-500/30 bg-emerald-500/10 p-4 text-center"><div className="text-3xl font-black text-emerald-400">{withAttorney.length}</div><div className="text-xs text-emerald-400/70">Has Attorney</div></div>
            <div className="rounded-2xl border border-amber-500/30 bg-amber-500/10 p-4 text-center"><div className="text-3xl font-black text-amber-400">{noAttorney.length}</div><div className="text-xs text-amber-400/70">No Attorney</div></div>
            <div className="rounded-2xl border border-red-500/30 bg-red-500/10 p-4 text-center"><div className="text-3xl font-black text-red-400">{errors.length}</div><div className="text-xs text-red-400/70">Errors</div></div>
          </div>
        )}

        {results.length > 0 && (
          <div className="rounded-2xl border border-slate-700 bg-slate-800/60 shadow-2xl">
            <div className="flex flex-wrap items-end justify-between gap-2 border-b border-slate-700 px-4 pt-4">
              <div className="flex gap-1">
                <button className={tabClass("all")} onClick={() => setActiveTab("all")}>All <span className="ml-1 rounded-full bg-slate-700 px-1.5 text-xs">{results.filter((r) => !r.pending && !r.error).length}</span></button>
                <button className={tabClass("attorney")} onClick={() => setActiveTab("attorney")}>✓ Has Attorney <span className="ml-1 rounded-full bg-slate-700 px-1.5 text-xs">{withAttorney.length}</span></button>
                <button className={tabClass("noAttorney")} onClick={() => setActiveTab("noAttorney")}>✗ No Attorney <span className="ml-1 rounded-full bg-slate-700 px-1.5 text-xs">{noAttorney.length}</span></button>
                {errors.length > 0 && <button className={tabClass("errors")} onClick={() => setActiveTab("errors")}>⚠ Errors <span className="ml-1 rounded-full bg-slate-700 px-1.5 text-xs">{errors.length}</span></button>}
              </div>
              <div className="flex flex-wrap gap-2 pb-1">
                <input value={searchFilter} onChange={(e) => setSearchFilter(e.target.value)} placeholder="Filter serial, owner, class, goods/services…" className="rounded-lg border border-slate-600 bg-slate-900 px-3 py-1.5 text-xs text-white placeholder-slate-500 focus:border-blue-500 focus:outline-none" />
                <button onClick={() => exportCSV(filteredRows, "trademark-results.csv")} className="rounded-lg bg-slate-700 px-3 py-1.5 text-xs font-medium hover:bg-slate-600">⬇ Export Current View</button>
                <button onClick={() => exportCSV(noAttorney, "no-attorney-prospects.csv")} className="rounded-lg bg-amber-600 px-3 py-1.5 text-xs font-medium hover:bg-amber-500">⬇ No-Attorney CSV</button>
              </div>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full min-w-[2200px] text-sm">
                <thead>
                  <tr className="border-b border-slate-700 text-left text-xs font-semibold uppercase tracking-wider text-slate-400">
                    {["Serial #","Mark","Class","Goods / Services","Filing Date","Owner","Entity","Status","Correspondent","Email","Attorney?","Attorney Name","Confidence","TSDR"].map((h) => <th key={h} className="whitespace-nowrap px-4 py-3">{h}</th>)}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-700/50">
                  {activeTab === "all" && results.filter((r) => r.pending).map((r) => (
                    <tr key={`${r.serialNumber}-pending`} className="animate-pulse"><td className="px-4 py-3 font-mono text-slate-500">{r.serialNumber}</td><td colSpan={13} className="px-4 py-3 text-slate-600">Looking up…</td></tr>
                  ))}

                  {filteredRows.map((r) => (
                    <tr key={r.serialNumber} className={`hover:bg-slate-700/30 ${r.error ? "bg-red-500/5" : r.hasAttorney ? "bg-emerald-500/5" : "bg-amber-500/5"}`}>
                      <td className="whitespace-nowrap px-4 py-3 font-mono text-xs text-slate-300">{r.serialNumber}{r.needsReview && <span className="ml-1 text-yellow-500" title="Needs review">⚑</span>}</td>
                      <td className="max-w-[220px] px-4 py-3 font-semibold text-white">{r.mark || "—"}</td>
                      <td className="whitespace-nowrap px-4 py-3 text-xs text-blue-300">{r.internationalClass || "—"}</td>
                      <td className="max-w-[420px] px-4 py-3 text-xs text-slate-300">{r.goodsServices || "—"}</td>
                      <td className="whitespace-nowrap px-4 py-3 text-xs text-slate-300">{r.filingDate || "—"}</td>
                      <td className="max-w-[300px] px-4 py-3 text-xs text-slate-200">{r.ownerName || "—"}</td>
                      <td className="max-w-[220px] px-4 py-3 text-xs text-slate-400">{r.entityType || "—"}</td>
                      <td className="max-w-[260px] px-4 py-3 text-xs text-slate-300" title={r.status}>{r.status || "—"}</td>
                      <td className="max-w-[280px] px-4 py-3 text-xs text-slate-300">{r.correspondentName || "—"}{r.phone && <div className="text-slate-500">{r.phone}</div>}</td>
                      <td className="max-w-[240px] px-4 py-3 text-xs">{r.email ? <a href={`mailto:${r.email}`} className="break-all text-blue-400 hover:underline">{r.email}</a> : "—"}</td>
                      <td className="whitespace-nowrap px-4 py-3"><Badge value={r.error ? null : r.hasAttorney} /></td>
                      <td className="max-w-[220px] px-4 py-3 text-xs text-slate-300">{r.attorneyName || "—"}</td>
                      <td className="whitespace-nowrap px-4 py-3">{r.error ? "—" : <ConfidencePill level={r.confidence} />}</td>
                      <td className="whitespace-nowrap px-4 py-3"><a href={getTsdrLink(r.serialNumber)} target="_blank" rel="noreferrer" className="text-xs text-blue-400 hover:underline">View →</a></td>
                    </tr>
                  ))}
                  {filteredRows.length === 0 && !isChecking && <tr><td colSpan={14} className="py-12 text-center text-slate-500">{searchFilter ? "No results match your filter." : "No results in this category yet."}</td></tr>}
                </tbody>
              </table>
            </div>

            <div className="border-t border-slate-700 px-4 py-3 text-xs text-slate-500">
              Showing {filteredRows.length} row(s). CSV export includes all detailed fields, including owner addresses, filing date, class, goods/services, mark description and disclaimer.
            </div>
          </div>
        )}

        {results.length === 0 && !isChecking && (
          <div className="rounded-2xl border border-dashed border-slate-700 p-16 text-center">
            <div className="mb-3 text-5xl">🔎</div>
            <p className="text-slate-400">Enter USPTO serial numbers and click <strong className="text-white">Lookup Trademarks</strong>.</p>
            <p className="mt-2 text-sm text-slate-500">Retrieves mark · owner · goods/services · class · filing date · status · contact · attorney status.</p>
          </div>
        )}

        <p className="mt-6 text-center text-xs text-slate-600">Data sourced from USPTO TSDR via your Google Apps Script proxy.</p>
      </div>
    </div>
  );
}
